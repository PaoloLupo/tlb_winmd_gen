use crate::idlgen::{EnumItemInfo, MethodInfo, TypeLibInfo};
use crossterm::{
    event::{self, DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEventKind},
    execute,
    terminal::{EnterAlternateScreen, LeaveAlternateScreen, disable_raw_mode, enable_raw_mode},
};
use ratatui::{
    Terminal,
    backend::{Backend, CrosstermBackend},
    layout::{Constraint, Direction, Layout},
    style::{Color, Modifier, Style},
    text::{Line, Span},
    widgets::{Block, Borders, Cell, List, ListItem, ListState, Paragraph, Row, Table, Wrap},
};
use std::{error::Error, io, path::PathBuf};

enum ViewMode {
    Idl,
    Structured,
}

struct App {
    tlb_path: PathBuf,
    type_lib_info: TypeLibInfo,
    types: Vec<(String, String)>,                 // Name, Kind
    filtered_types: Vec<(usize, String, String)>, // Original Index, Name, Kind
    list_state: ListState,
    current_idl: String,
    current_methods: Vec<MethodInfo>,
    current_enums: Vec<EnumItemInfo>,
    search_query: String,
    view_mode: ViewMode,
}

impl App {
    fn new(tlb_path: PathBuf) -> Result<Self, Box<dyn Error>> {
        let mut type_lib_info = TypeLibInfo::new();
        type_lib_info.load_type_lib(&tlb_path)?;

        let count = type_lib_info.get_type_info_count();
        let mut types = Vec::new();
        for i in 0..count {
            if let Ok((name, kind)) = type_lib_info.get_type_name_and_kind(i) {
                types.push((name, kind));
            }
        }

        let mut app = App {
            tlb_path,
            type_lib_info,
            types,
            filtered_types: Vec::new(),
            list_state: ListState::default(),
            current_idl: String::new(),
            current_methods: Vec::new(),
            current_enums: Vec::new(),
            search_query: String::new(),
            view_mode: ViewMode::Structured,
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

        if !self.filtered_types.is_empty() {
            self.list_state.select(Some(0));
        } else {
            self.list_state.select(None);
        }
        self.update_selection();
    }

    fn update_selection(&mut self) {
        if let Some(selected_idx) = self.list_state.selected() {
            if let Some((original_idx, _, _)) = self.filtered_types.get(selected_idx) {
                let idx = *original_idx as u32;
                if let Ok(idl) = self.type_lib_info.get_type_idl(idx) {
                    self.current_idl = idl;
                } else {
                    self.current_idl = "Error loading IDL".to_string();
                }

                if let Ok(methods) = self.type_lib_info.get_type_methods(idx) {
                    self.current_methods = methods;
                } else {
                    self.current_methods = Vec::new();
                }

                if let Ok(enums) = self.type_lib_info.get_type_enums(idx) {
                    self.current_enums = enums;
                } else {
                    self.current_enums = Vec::new();
                }
            } else {
                self.current_idl = String::new();
                self.current_methods = Vec::new();
                self.current_enums = Vec::new();
            }
        } else {
            self.current_idl = String::new();
            self.current_methods = Vec::new();
            self.current_enums = Vec::new();
        }
    }

    fn next(&mut self) {
        if self.filtered_types.is_empty() {
            return;
        }
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
        self.update_selection();
    }

    fn previous(&mut self) {
        if self.filtered_types.is_empty() {
            return;
        }
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
        self.update_selection();
    }

    fn toggle_view(&mut self) {
        self.view_mode = match self.view_mode {
            ViewMode::Idl => ViewMode::Structured,
            ViewMode::Structured => ViewMode::Idl,
        };
    }
}

pub fn run(tlb_path: PathBuf) -> Result<(), Box<dyn Error>> {
    // Setup terminal
    enable_raw_mode()?;
    let mut stdout = io::stdout();
    execute!(stdout, EnterAlternateScreen, EnableMouseCapture)?;
    let backend = CrosstermBackend::new(stdout);
    let mut terminal = Terminal::new(backend)?;

    // Create app
    let app = App::new(tlb_path);

    let res = match app {
        Ok(mut app) => run_app(&mut terminal, &mut app),
        Err(e) => Err(e),
    };

    // Restore terminal
    disable_raw_mode()?;
    execute!(
        terminal.backend_mut(),
        LeaveAlternateScreen,
        DisableMouseCapture
    )?;
    terminal.show_cursor()?;

    if let Err(err) = res {
        println!("{:?}", err);
    }

    Ok(())
}

fn run_app<B: Backend>(terminal: &mut Terminal<B>, app: &mut App) -> Result<(), Box<dyn Error>> {
    loop {
        terminal.draw(|f| ui(f, app))?;

        if let Event::Key(key) = event::read()? {
            if key.kind == KeyEventKind::Press {
                match key.code {
                    KeyCode::Char('q') if app.search_query.is_empty() => return Ok(()),
                    KeyCode::Esc => return Ok(()),
                    KeyCode::Down => app.next(),
                    KeyCode::Up => app.previous(),
                    KeyCode::Tab => app.toggle_view(),
                    KeyCode::Char(c) => {
                        app.search_query.push(c);
                        app.update_filter();
                    }
                    KeyCode::Backspace => {
                        app.search_query.pop();
                        app.update_filter();
                    }
                    _ => {}
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

    // Search bar
    let search_paragraph = Paragraph::new(app.search_query.as_str())
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title("Search (Tab to toggle view)"),
        )
        .style(Style::default().fg(Color::White));
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
        .block(
            Block::default()
                .borders(Borders::ALL)
                .title(format!("Types - {}", app.tlb_path.display())),
        )
        .highlight_style(
            Style::default()
                .bg(Color::Blue)
                .fg(Color::White)
                .add_modifier(Modifier::BOLD),
        )
        .highlight_symbol(">> ");

    f.render_stateful_widget(list, content_chunks[0], &mut app.list_state);

    // Right panel
    match app.view_mode {
        ViewMode::Idl => {
            let idl_paragraph = Paragraph::new(app.current_idl.as_str())
                .block(Block::default().borders(Borders::ALL).title("IDL Preview"))
                .wrap(Wrap { trim: false });
            f.render_widget(idl_paragraph, content_chunks[1]);
        }
        ViewMode::Structured => {
            if !app.current_methods.is_empty() {
                let header_cells = ["Method Signature"]
                    .iter()
                    .map(|h| Cell::from(*h).style(Style::default().fg(Color::Yellow)));
                let header = Row::new(header_cells)
                    .style(Style::default().bg(Color::DarkGray))
                    .height(1);

                let rows = app.current_methods.iter().map(|method| {
                    // Format: ƒ "name" (→ type param, ...) -> ret_type
                    let mut spans = Vec::new();
                    spans.push(Span::styled("ƒ ", Style::default().fg(Color::Magenta)));
                    spans.push(Span::styled(
                        format!("\"{}\" ", method.name),
                        Style::default().fg(Color::Cyan),
                    ));
                    spans.push(Span::raw("("));

                    for (i, param) in method.params.iter().enumerate() {
                        if i > 0 {
                            spans.push(Span::raw(", "));
                        }

                        // Icons for flags
                        if param.flags.contains(&"in".to_string()) {
                            spans.push(Span::styled("→ ", Style::default().fg(Color::Green)));
                        }
                        if param.flags.contains(&"out".to_string()) {
                            spans.push(Span::styled("← ", Style::default().fg(Color::Red)));
                        }
                        if param.flags.contains(&"defaultvalue".to_string()) {
                            spans.push(Span::styled("* ", Style::default().fg(Color::Blue)));
                        }
                        if param.flags.contains(&"optional".to_string()) {
                            spans.push(Span::styled("? ", Style::default().fg(Color::Yellow)));
                        }

                        spans.push(Span::styled(
                            format!("{} ", param.type_name),
                            Style::default().fg(Color::White),
                        ));
                        spans.push(Span::raw(&param.name));
                    }

                    spans.push(Span::raw(") -> "));
                    spans.push(Span::styled(
                        format!("{}", method.ret_type),
                        Style::default().fg(Color::Green),
                    ));

                    Row::new(vec![Cell::from(Line::from(spans))])
                });

                let table = Table::new(rows, [Constraint::Percentage(100)])
                    .header(header)
                    .block(Block::default().borders(Borders::ALL).title("Methods"));

                f.render_widget(table, content_chunks[1]);
            } else if !app.current_enums.is_empty() {
                let header_cells = ["Name", "Value"]
                    .iter()
                    .map(|h| Cell::from(*h).style(Style::default().fg(Color::Yellow)));
                let header = Row::new(header_cells)
                    .style(Style::default().bg(Color::DarkGray))
                    .height(1);

                let rows = app.current_enums.iter().map(|item| {
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
                .block(Block::default().borders(Borders::ALL).title("Enum Values"));

                f.render_widget(table, content_chunks[1]);
            } else {
                // Fallback to IDL if no structured data available
                let idl_paragraph = Paragraph::new(app.current_idl.as_str())
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .title("IDL Preview (No structured data)"),
                    )
                    .wrap(Wrap { trim: false });
                f.render_widget(idl_paragraph, content_chunks[1]);
            }
        }
    }
}
