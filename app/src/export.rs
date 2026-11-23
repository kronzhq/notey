//! Exports the current buffer to PDF or DOCX (not the existing .txt/.md save path).
//! Markdown is rendered like the preview; anything else is exported as the editor shows it.
//! Both use the Settings font; markdown code uses Lilex, like the preview's.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use anyhow::{Context as _, Result};
use docx_rs::{
    AlignmentType, BorderType, BreakType, Docx, Hyperlink, HyperlinkType, LineSpacing,
    LineSpacingType, Paragraph, Run, RunFonts, Shading, ShdType, Table, TableBorder, TableBorders,
    TableCell, TableRow, WidthType,
};
use markdown::{
    ParseOptions,
    mdast::{AlignKind, Node},
};
use printpdf::{
    Color, FontId, Line, LinePoint, Mm, Op, PaintMode, ParsedFont, PdfDocument, PdfFontHandle,
    PdfPage, PdfSaveOptions, Point, Polygon, PolygonRing, Pt, Rgb, TextItem, WindingOrder,
};

use crate::{LILEX_BOLD, LILEX_REGULAR, LINE_SPACING};

#[derive(Clone, Copy, PartialEq, Debug)]
pub enum ExportFormat {
    Pdf,
    Docx,
}

impl ExportFormat {
    pub fn extension(self) -> &'static str {
        match self {
            ExportFormat::Pdf => "pdf",
            ExportFormat::Docx => "docx",
        }
    }

    /// `family`/`size` are the Settings font, which the editor and preview both use.
    pub fn write(
        self,
        text: &str,
        markdown: bool,
        family: &str,
        size: f32,
        path: &Path,
    ) -> Result<()> {
        let blocks = if markdown {
            markdown_blocks(text)
        } else {
            vec![Block::plain_code(text)]
        };
        let code_family = if markdown { "Lilex" } else { family };
        match self {
            ExportFormat::Pdf => write_pdf(&blocks, family, code_family, size, path),
            ExportFormat::Docx => write_docx(&blocks, family, code_family, size, path),
        }
    }
}

// ---- Document model shared by both writers ----

#[derive(Clone, Copy, PartialEq, Debug)]
enum Align {
    Left,
    Center,
    Right,
}

#[derive(Clone, Copy, PartialEq, Debug)]
enum Kind {
    Heading(u8),
    Para,
    /// Fenced/indented markdown code block (shaded box).
    Code,
    /// Raw plain-text document export (no shaded box).
    PlainCode,
    Rule,
    Table,
}

#[derive(Clone, Default, PartialEq, Debug)]
struct Style {
    bold: bool,
    italic: bool,
    strike: bool,
    code: bool,
    muted: bool,
    link: Option<String>,
}

#[derive(Clone, Debug)]
struct Span {
    text: String,
    style: Style,
}

#[derive(Debug)]
struct TableData {
    align: Vec<Align>,
    /// Row 0 is the header row; subsequent rows are body rows.
    rows: Vec<Vec<Vec<Span>>>,
}

#[derive(Debug)]
struct Block {
    kind: Kind,
    /// List nesting depth.
    indent: u8,
    /// Blockquote nesting depth (rendered with left accent bar(s)).
    quote_depth: u8,
    /// Task-list checkbox state on the first line of a list item.
    checkbox: Option<bool>,
    spans: Vec<Span>,
    table: Option<TableData>,
}

impl Block {
    fn code(text: &str, indent: u8, quote_depth: u8) -> Self {
        Block {
            kind: Kind::Code,
            indent,
            quote_depth,
            checkbox: None,
            spans: vec![Span {
                text: text.to_string(),
                style: Style {
                    code: true,
                    ..Style::default()
                },
            }],
            table: None,
        }
    }

    fn plain_code(text: &str) -> Self {
        Block {
            kind: Kind::PlainCode,
            ..Self::code(text, 0, 0)
        }
    }
}

fn markdown_blocks(text: &str) -> Vec<Block> {
    let mut out = Vec::new();
    match markdown::to_mdast(text, &ParseOptions::gfm()) {
        Ok(root) => {
            let mut defs = HashMap::new();
            collect_definitions(&root, &mut defs);
            let mut footnotes = Vec::new();
            collect_blocks(&root, 0, 0, &defs, &mut footnotes, &mut out);
            if !footnotes.is_empty() {
                out.push(Block {
                    kind: Kind::Rule,
                    indent: 0,
                    quote_depth: 0,
                    checkbox: None,
                    spans: vec![],
                    table: None,
                });
                for (label, mut blocks) in footnotes {
                    let prefix = Span {
                        text: format!("[{label}] "),
                        style: Style {
                            bold: true,
                            muted: true,
                            ..Style::default()
                        },
                    };
                    if let Some(first) = blocks.first_mut()
                        && matches!(first.kind, Kind::Para | Kind::Heading(_))
                    {
                        first.spans.insert(0, prefix);
                    } else {
                        out.push(Block {
                            kind: Kind::Para,
                            indent: 0,
                            quote_depth: 0,
                            checkbox: None,
                            spans: vec![prefix],
                            table: None,
                        });
                    }
                    out.extend(blocks);
                }
            }
        }
        // Plain markdown never fails to parse (only MDX can); fall back to raw text anyway.
        Err(_) => out.push(Block::plain_code(text)),
    }
    out
}

fn collect_definitions(node: &Node, defs: &mut HashMap<String, String>) {
    if let Node::Definition(d) = node {
        defs.insert(d.identifier.to_lowercase(), d.url.clone());
    }
    if let Some(children) = node.children() {
        for child in children {
            collect_definitions(child, defs);
        }
    }
}

fn collect_blocks(
    node: &Node,
    indent: u8,
    quote_depth: u8,
    defs: &HashMap<String, String>,
    footnotes: &mut Vec<(String, Vec<Block>)>,
    out: &mut Vec<Block>,
) {
    let push_inline = |kind: Kind, children: &[Node], out: &mut Vec<Block>| {
        let mut spans = Vec::new();
        let base_style = Style {
            muted: quote_depth > 0 || kind == Kind::Heading(6),
            ..Style::default()
        };
        collect_inline(children, &base_style, defs, &mut spans);
        out.push(Block {
            kind,
            indent,
            quote_depth,
            checkbox: None,
            spans,
            table: None,
        });
    };
    match node {
        Node::Heading(h) => push_inline(Kind::Heading(h.depth), &h.children, out),
        Node::Paragraph(p) => push_inline(Kind::Para, &p.children, out),
        Node::Code(c) => out.push(Block::code(&c.value, indent, quote_depth)),
        Node::ThematicBreak(_) => out.push(Block {
            kind: Kind::Rule,
            indent,
            quote_depth,
            checkbox: None,
            spans: vec![],
            table: None,
        }),
        Node::Blockquote(q) => {
            for c in &q.children {
                collect_blocks(c, indent, quote_depth + 1, defs, footnotes, out);
            }
        }
        Node::List(list) => {
            for (i, item) in list.children.iter().enumerate() {
                let checked = match item {
                    Node::ListItem(li) => li.checked,
                    _ => None,
                };
                let marker = if checked.is_some() {
                    None
                } else if list.ordered {
                    Some(format!("{}. ", list.start.unwrap_or(1) as usize + i))
                } else if indent == 0 {
                    Some("\u{2022} ".to_string())
                } else {
                    Some("- ".to_string())
                };
                let first = out.len();
                for child in item.children().into_iter().flatten() {
                    collect_blocks(child, indent + 1, quote_depth, defs, footnotes, out);
                }
                if out.len() == first {
                    out.push(Block {
                        kind: Kind::Para,
                        indent: indent + 1,
                        quote_depth,
                        checkbox: checked,
                        spans: marker
                            .map(|m| {
                                vec![Span {
                                    text: m,
                                    style: Style {
                                        muted: quote_depth > 0,
                                        ..Style::default()
                                    },
                                }]
                            })
                            .unwrap_or_default(),
                        table: None,
                    });
                } else if let Some(block) = out.get_mut(first)
                    && matches!(block.kind, Kind::Para | Kind::Heading(_))
                {
                    if let Some(c) = checked {
                        block.checkbox = Some(c);
                    } else if let Some(m) = marker {
                        block.spans.insert(
                            0,
                            Span {
                                text: m,
                                style: Style {
                                    muted: quote_depth > 0,
                                    ..Style::default()
                                },
                            },
                        );
                    }
                }
            }
        }
        Node::Table(t) => {
            let align: Vec<Align> = t
                .align
                .iter()
                .map(|a| match a {
                    AlignKind::Center => Align::Center,
                    AlignKind::Right => Align::Right,
                    AlignKind::Left | AlignKind::None => Align::Left,
                })
                .collect();
            let mut rows = Vec::new();
            for row in &t.children {
                let mut cells = Vec::new();
                for cell in row.children().into_iter().flatten() {
                    let mut spans = Vec::new();
                    let base_style = Style {
                        muted: quote_depth > 0,
                        ..Style::default()
                    };
                    collect_inline(
                        cell.children().map_or(&[][..], |c| c),
                        &base_style,
                        defs,
                        &mut spans,
                    );
                    cells.push(spans);
                }
                rows.push(cells);
            }
            if !rows.is_empty() {
                out.push(Block {
                    kind: Kind::Table,
                    indent,
                    quote_depth,
                    checkbox: None,
                    spans: vec![],
                    table: Some(TableData { align, rows }),
                });
            }
        }
        Node::FootnoteDefinition(f) => {
            let label = f.label.as_deref().unwrap_or(&f.identifier).to_string();
            let mut fn_blocks = Vec::new();
            for c in &f.children {
                collect_blocks(c, 0, 0, defs, footnotes, &mut fn_blocks);
            }
            footnotes.push((label, fn_blocks));
        }
        Node::Definition(_) => {}
        Node::Html(h) => {
            let trimmed = h.value.trim().to_ascii_lowercase();
            if matches!(trimmed.as_str(), "<hr>" | "<hr/>" | "<hr />") {
                out.push(Block {
                    kind: Kind::Rule,
                    indent,
                    quote_depth,
                    checkbox: None,
                    spans: vec![],
                    table: None,
                });
            }
        }
        other => {
            for c in other.children().into_iter().flatten() {
                collect_blocks(c, indent, quote_depth, defs, footnotes, out);
            }
        }
    }
}

fn collect_inline(
    nodes: &[Node],
    style: &Style,
    defs: &HashMap<String, String>,
    out: &mut Vec<Span>,
) {
    for node in nodes {
        match node {
            Node::Text(t) => out.push(Span {
                text: t.value.replace('\n', " "),
                style: style.clone(),
            }),
            Node::InlineCode(c) => out.push(Span {
                text: c.value.clone(),
                style: Style {
                    code: true,
                    ..style.clone()
                },
            }),
            Node::Break(_) => out.push(Span {
                text: "\n".into(),
                style: style.clone(),
            }),
            Node::Strong(s) => collect_inline(
                &s.children,
                &Style {
                    bold: true,
                    ..style.clone()
                },
                defs,
                out,
            ),
            Node::Emphasis(e) => collect_inline(
                &e.children,
                &Style {
                    italic: true,
                    ..style.clone()
                },
                defs,
                out,
            ),
            Node::Delete(d) => collect_inline(
                &d.children,
                &Style {
                    strike: true,
                    ..style.clone()
                },
                defs,
                out,
            ),
            Node::Link(l) => {
                let link_style = Style {
                    link: Some(l.url.clone()),
                    ..style.clone()
                };
                if l.children.is_empty() {
                    out.push(Span {
                        text: l.url.clone(),
                        style: link_style,
                    });
                } else {
                    collect_inline(&l.children, &link_style, defs, out);
                }
            }
            Node::LinkReference(lr) => {
                let url = defs.get(&lr.identifier.to_lowercase()).cloned();
                let link_style = Style {
                    link: url.or_else(|| style.link.clone()),
                    ..style.clone()
                };
                collect_inline(&lr.children, &link_style, defs, out);
            }
            Node::Image(i) => {
                let label = if i.alt.is_empty() {
                    "[Image]".to_string()
                } else {
                    format!("[Image: {}]", i.alt)
                };
                out.push(Span {
                    text: label,
                    style: Style {
                        italic: true,
                        muted: true,
                        ..style.clone()
                    },
                });
            }
            Node::ImageReference(ir) => {
                let label = if ir.alt.is_empty() {
                    "[Image]".to_string()
                } else {
                    format!("[Image: {}]", ir.alt)
                };
                out.push(Span {
                    text: label,
                    style: Style {
                        italic: true,
                        muted: true,
                        ..style.clone()
                    },
                });
            }
            Node::FootnoteReference(fr) => {
                let label = fr.label.as_deref().unwrap_or(&fr.identifier);
                out.push(Span {
                    text: format!("[{label}]"),
                    style: Style {
                        muted: true,
                        ..style.clone()
                    },
                });
            }
            Node::Html(h) => {
                let trimmed = h.value.trim().to_ascii_lowercase();
                if matches!(trimmed.as_str(), "<br>" | "<br/>" | "<br />") {
                    out.push(Span {
                        text: "\n".into(),
                        style: style.clone(),
                    });
                }
            }
            other => collect_inline(other.children().map_or(&[][..], |c| c), style, defs, out),
        }
    }
}

// ---- PDF ----

const PAGE_W_MM: f32 = 210.0;
const PAGE_H_MM: f32 = 297.0;
const MARGIN_MM: f32 = 20.0;
const INDENT_PT: f32 = 18.0;
const QUOTE_INDENT_PT: f32 = 14.0;

/// H1–H6 sizes relative to the body text (1.125 ratio), shared with the preview.
pub const HEADING_SCALE: [f32; 6] = [1.602, 1.424, 1.266, 1.125, 1.0, 0.889];
pub const CODE_SCALE: f32 = 0.889;

fn block_size(kind: Kind, body: f32) -> f32 {
    body * match kind {
        Kind::Heading(level) => HEADING_SCALE[level as usize - 1],
        Kind::Code => CODE_SCALE,
        _ => 1.0,
    }
}

fn line_height(kind: Kind) -> f32 {
    match kind {
        Kind::Heading(_) => 1.25,
        _ => LINE_SPACING,
    }
}

/// Space above a block, for both PDF and DOCX.
fn gap_before(prev: Option<&Block>, block: &Block, body: f32) -> f32 {
    let Some(prev) = prev else { return 0.0 };
    body * match (prev.kind, block.kind) {
        (Kind::Heading(_), Kind::Heading(_)) => 0.75,
        (_, Kind::Heading(1 | 2)) => 1.75,
        (_, Kind::Heading(_)) => 1.25,
        (Kind::Heading(_), _) => 0.5,
        (Kind::Para, Kind::Para) if prev.indent > 0 && block.indent > 0 => 0.25,
        _ => 0.75,
    }
}

fn rgb(r: f32, g: f32, b: f32) -> Color {
    Color::Rgb(Rgb {
        r,
        g,
        b,
        icc_profile: None,
    })
}

fn fill_rect(ops: &mut Vec<Op>, x: f32, y_bottom: f32, w: f32, h: f32, color: &Color) {
    if w <= 0.0 || h <= 0.0 {
        return;
    }
    let pts = [
        (x, y_bottom),
        (x + w, y_bottom),
        (x + w, y_bottom + h),
        (x, y_bottom + h),
    ]
    .map(|(px, py)| LinePoint {
        p: Point {
            x: Pt(px),
            y: Pt(py),
        },
        bezier: false,
    })
    .to_vec();
    ops.push(Op::SetFillColor { col: color.clone() });
    ops.push(Op::DrawPolygon {
        polygon: Polygon {
            rings: vec![PolygonRing { points: pts }],
            mode: PaintMode::Fill,
            winding_order: WindingOrder::NonZero,
        },
    });
}

fn draw_line(ops: &mut Vec<Op>, x1: f32, y1: f32, x2: f32, y2: f32, thickness: f32, color: &Color) {
    ops.push(Op::SetOutlineColor { col: color.clone() });
    ops.push(Op::SetOutlineThickness { pt: Pt(thickness) });
    ops.push(Op::DrawLine {
        line: Line {
            points: vec![
                LinePoint {
                    p: Point {
                        x: Pt(x1),
                        y: Pt(y1),
                    },
                    bezier: false,
                },
                LinePoint {
                    p: Point {
                        x: Pt(x2),
                        y: Pt(y2),
                    },
                    bezier: false,
                },
            ],
            is_closed: false,
        },
    });
}

fn stroke_rect(
    ops: &mut Vec<Op>,
    x: f32,
    y_bottom: f32,
    w: f32,
    h: f32,
    thickness: f32,
    color: &Color,
) {
    draw_line(ops, x, y_bottom, x + w, y_bottom, thickness, color);
    draw_line(ops, x, y_bottom + h, x + w, y_bottom + h, thickness, color);
    draw_line(ops, x, y_bottom, x, y_bottom + h, thickness, color);
    draw_line(ops, x + w, y_bottom, x + w, y_bottom + h, thickness, color);
}

fn draw_quote_bars(
    ops: &mut Vec<Op>,
    base_x: f32,
    quote_depth: u8,
    y_bottom: f32,
    height: f32,
    color: &Color,
) {
    for q in 0..quote_depth {
        let bar_x = base_x + q as f32 * QUOTE_INDENT_PT + 3.0;
        fill_rect(ops, bar_x, y_bottom, 2.2, height, color);
    }
}

fn draw_checkbox(
    ops: &mut Vec<Op>,
    x: f32,
    y_baseline: f32,
    size: f32,
    checked: bool,
    border: &Color,
    accent: &Color,
) {
    let box_sz = size * 0.72;
    let y0 = y_baseline + size * 0.04;
    if checked {
        fill_rect(ops, x, y0, box_sz, box_sz, accent);
        let white = rgb(1.0, 1.0, 1.0);
        let p1 = (x + box_sz * 0.20, y0 + box_sz * 0.48);
        let p2 = (x + box_sz * 0.42, y0 + box_sz * 0.24);
        let p3 = (x + box_sz * 0.80, y0 + box_sz * 0.74);
        draw_line(ops, p1.0, p1.1, p2.0, p2.1, 1.1, &white);
        draw_line(ops, p2.0, p2.1, p3.0, p3.1, 1.1, &white);
    } else {
        stroke_rect(ops, x, y0, box_sz, box_sz, 0.75, border);
    }
}

/// Where a font face comes from: the bundled Lilex, or a system font file (+ face index).
#[derive(Clone, PartialEq, Debug)]
enum Source {
    Lilex { bold: bool },
    File(PathBuf, usize),
}

/// Finds `family`'s face with exactly these traits via CoreText; None if it has no such face.
fn system_face(family: &str, bold: bool, italic: bool) -> Option<Source> {
    use core_foundation::{array::CFArray, base::TCFType};
    use core_text::font_descriptor::{
        CTFontDescriptor, SymbolicTraitAccessors as _, kCTFontBoldTrait, kCTFontItalicTrait,
    };

    let descs = core_text::font_collection::create_for_family(family)?.get_descriptors()?;
    let base = core_text::font::new_from_descriptor(&*descs.get(0)?, 12.0);
    let want =
        (if bold { kCTFontBoldTrait } else { 0 }) | (if italic { kCTFontItalicTrait } else { 0 });
    let font = base.clone_with_symbolic_traits(want, kCTFontBoldTrait | kCTFontItalicTrait)?;
    let traits = font.symbolic_traits();
    if traits.is_bold() != bold || traits.is_italic() != italic || font.family_name() != family {
        return None;
    }
    let url = font.url()?;
    // A .ttc holds several faces; CoreText lists them in file order.
    let name = font.postscript_name();
    let faces = unsafe {
        core_text::font_manager::CTFontManagerCreateFontDescriptorsFromURL(
            url.as_concrete_TypeRef(),
        )
    };
    let index = (!faces.is_null())
        .then(|| unsafe { CFArray::<CTFontDescriptor>::wrap_under_create_rule(faces) })
        .and_then(|faces| faces.iter().position(|d| d.font_name() == name))
        .unwrap_or(0);
    // ponytail: a variable font (e.g. SF Pro) embeds its default instance, so its bold
    // prints regular; add an instancer if that matters.
    Some(Source::File(url.to_path()?, index))
}

/// Regular, bold, italic, bold italic for `family`, falling back per face to the closest one
/// that exists, and to Lilex when the family isn't a system font (Lilex is app-bundled).
fn family_faces(family: &str) -> [Source; 4] {
    let lilex = |bold| Source::Lilex { bold };
    let Some(regular) = (family != "Lilex")
        .then(|| system_face(family, false, false))
        .flatten()
    else {
        return [lilex(false), lilex(true), lilex(false), lilex(true)];
    };
    let bold = system_face(family, true, false).unwrap_or(regular.clone());
    let italic = system_face(family, false, true).unwrap_or(regular.clone());
    let bold_italic = system_face(family, true, true).unwrap_or(bold.clone());
    [regular, bold, italic, bold_italic]
}

/// The document's embedded fonts, each also parsed for glyph widths so text wraps correctly.
struct Fonts {
    faces: Vec<(FontId, ParsedFont)>,
    /// Indexes into `faces`: regular, bold, italic, bold italic, then code.
    slots: [usize; 5],
}

impl Fonts {
    fn load(doc: &mut PdfDocument, family: &str, code_family: &str) -> Result<Self> {
        let [r, b, i, bi] = family_faces(family);
        let code = family_faces(code_family)[0].clone();
        let mut sources: Vec<Source> = vec![];
        let mut faces = vec![];
        let slots = [r, b, i, bi, code].map(|src| {
            sources.iter().position(|s| *s == src).unwrap_or_else(|| {
                sources.push(src);
                sources.len() - 1
            })
        });
        for src in &sources {
            let (bytes, index) = match src {
                Source::Lilex { bold: false } => (LILEX_REGULAR.to_vec(), 0),
                Source::Lilex { bold: true } => (LILEX_BOLD.to_vec(), 0),
                Source::File(path, index) => (
                    std::fs::read(path).with_context(|| format!("reading {}", path.display()))?,
                    *index,
                ),
            };
            let parsed = ParsedFont::from_bytes(&bytes, index, &mut Vec::new())
                .with_context(|| format!("parsing font {src:?}"))?;
            faces.push((doc.add_font(&parsed), parsed));
        }
        Ok(Fonts { faces, slots })
    }

    fn face(&self, style: &Style) -> usize {
        match (style.code, style.bold, style.italic) {
            (true, _, _) => self.slots[4],
            (_, bold, italic) => self.slots[bold as usize + 2 * italic as usize],
        }
    }

    fn width(&self, face: usize, size: f32, text: &str) -> f32 {
        let parsed = &self.faces[face].1;
        let upm = parsed.font_metrics.units_per_em.max(1) as f32;
        text.chars()
            .map(|c| {
                parsed.lookup_glyph_index(c as u32).map_or(size * 0.6, |g| {
                    parsed.get_horizontal_advance(g) as f32 / upm * size
                })
            })
            .sum()
    }
}

#[derive(Clone, Debug)]
struct PdfRun {
    font: usize,
    style: Style,
    text: String,
}

type PdfLine = Vec<PdfRun>;

/// Greedy word wrap over styled spans; code keeps its own line breaks and hard-wraps.
fn wrap_spans(
    spans: &[Span],
    is_code: bool,
    force_bold: bool,
    size: f32,
    max_w: f32,
    m: &Fonts,
) -> Vec<PdfLine> {
    let max_w = max_w.max(size);
    let mut lines: Vec<PdfLine> = vec![vec![]];
    let mut width = 0.0;
    for span in spans {
        let effective = Style {
            bold: force_bold || span.style.bold,
            ..span.style.clone()
        };
        let font = m.face(&effective);
        let tokens: Vec<&str> = if is_code {
            span.text.split_inclusive('\n').collect()
        } else {
            span.text.split_inclusive([' ', '\n']).collect()
        };
        for token in tokens {
            let (mut rest, newline) = match token.strip_suffix('\n') {
                Some(t) => (t, true),
                None => (token, false),
            };
            while !rest.is_empty() {
                // Trailing spaces may hang past the margin; they're invisible.
                if width + m.width(font, size, rest.trim_end()) <= max_w {
                    push_run(lines.last_mut().unwrap(), font, &effective, rest);
                    width += m.width(font, size, rest);
                    break;
                }
                if width > 0.0 {
                    lines.push(vec![]);
                    width = 0.0;
                    continue;
                }
                // Alone on the line and still too wide: hard-break where it overflows.
                let mut acc = 0.0;
                let cut = rest
                    .char_indices()
                    .find(|&(i, c)| {
                        acc += m.width(font, size, c.encode_utf8(&mut [0; 4]));
                        i > 0 && acc > max_w
                    })
                    .map_or(rest.len(), |(i, _)| i);
                push_run(lines.last_mut().unwrap(), font, &effective, &rest[..cut]);
                rest = &rest[cut..];
                lines.push(vec![]);
            }
            if newline {
                lines.push(vec![]);
                width = 0.0;
            }
        }
    }
    if is_code && lines.len() > 1 && lines.last().is_some_and(Vec::is_empty) {
        lines.pop(); // trailing newline of the buffer
    }
    lines
}

fn wrap_block(block: &Block, size: f32, max_w: f32, m: &Fonts) -> Vec<PdfLine> {
    let is_code = matches!(block.kind, Kind::Code | Kind::PlainCode);
    let force_bold = matches!(block.kind, Kind::Heading(_));
    wrap_spans(&block.spans, is_code, force_bold, size, max_w, m)
}

fn push_run(line: &mut PdfLine, font: usize, style: &Style, text: &str) {
    match line.last_mut() {
        Some(r) if r.font == font && r.style == *style => r.text.push_str(text),
        _ => line.push(PdfRun {
            font,
            style: style.clone(),
            text: text.to_string(),
        }),
    }
}

fn line_width(line: &PdfLine, size: f32, m: &Fonts) -> f32 {
    let mut w = 0.0;
    for (i, run) in line.iter().enumerate() {
        let text = if i + 1 == line.len() {
            run.text.trim_end()
        } else {
            run.text.as_str()
        };
        w += m.width(run.font, size, text);
    }
    w
}

struct PdfPalette {
    text: Color,
    muted: Color,
    link: Color,
    border: Color,
    quote_bar: Color,
    code_bg: Color,
    inline_code_bg: Color,
    table_header_bg: Color,
    table_stripe_bg: Color,
}

impl PdfPalette {
    fn standard() -> Self {
        Self {
            text: rgb(0.08, 0.09, 0.11),
            muted: rgb(0.38, 0.42, 0.48),
            link: rgb(0.04, 0.41, 0.85),
            border: rgb(0.80, 0.83, 0.86),
            quote_bar: rgb(0.75, 0.78, 0.82),
            code_bg: rgb(0.95, 0.96, 0.97),
            inline_code_bg: rgb(0.93, 0.94, 0.96),
            table_header_bg: rgb(0.93, 0.95, 0.97),
            table_stripe_bg: rgb(0.98, 0.985, 0.99),
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn draw_styled_line(
    ops: &mut Vec<Op>,
    line: &PdfLine,
    x: f32,
    y_baseline: f32,
    size: f32,
    inline_code_pill: bool,
    palette: &PdfPalette,
    m: &Fonts,
) {
    let mut cur_x = x;
    for (i, run) in line.iter().enumerate() {
        let draw_text = if i + 1 == line.len() {
            run.text.trim_end()
        } else {
            run.text.as_str()
        };
        let run_w = m.width(run.font, size, draw_text);
        if inline_code_pill && run.style.code && run_w > 0.0 {
            fill_rect(
                ops,
                cur_x - 1.5,
                y_baseline - size * 0.22,
                run_w + 3.0,
                size * 1.12,
                &palette.inline_code_bg,
            );
        }
        let color = if run.style.link.is_some() {
            &palette.link
        } else if run.style.muted {
            &palette.muted
        } else {
            &palette.text
        };
        if !draw_text.is_empty() {
            ops.push(Op::StartTextSection);
            ops.push(Op::SetTextCursor {
                pos: Point {
                    x: Pt(cur_x),
                    y: Pt(y_baseline),
                },
            });
            ops.push(Op::SetFillColor { col: color.clone() });
            ops.push(Op::SetFont {
                font: PdfFontHandle::External(m.faces[run.font].0.clone()),
                size: Pt(size),
            });
            ops.push(Op::ShowText {
                items: vec![TextItem::Text(draw_text.to_string())],
            });
            ops.push(Op::EndTextSection);
        }
        if run.style.link.is_some() && run_w > 0.0 {
            draw_line(
                ops,
                cur_x,
                y_baseline - size * 0.14,
                cur_x + run_w,
                y_baseline - size * 0.14,
                0.5,
                &palette.link,
            );
        }
        if run.style.strike && run_w > 0.0 {
            draw_line(
                ops,
                cur_x,
                y_baseline + size * 0.30,
                cur_x + run_w,
                y_baseline + size * 0.30,
                0.6,
                color,
            );
        }
        cur_x += m.width(run.font, size, &run.text);
    }
}

fn compute_table_col_widths(
    table: &TableData,
    ncols: usize,
    size: f32,
    pad_x: f32,
    avail_w: f32,
    m: &Fonts,
) -> Vec<f32> {
    let mut natural = vec![0.0_f32; ncols];
    for (r_idx, row) in table.rows.iter().enumerate() {
        let force_bold = r_idx == 0;
        for (c_idx, cell) in row.iter().enumerate().take(ncols) {
            let mut w = 0.0;
            for span in cell {
                let effective = Style {
                    bold: force_bold || span.style.bold,
                    ..span.style.clone()
                };
                let face = m.face(&effective);
                w += m.width(face, size, &span.text);
            }
            natural[c_idx] = natural[c_idx].max(w + pad_x * 2.0 + 4.0);
        }
    }
    let min_w = (avail_w / ncols as f32).clamp(20.0, 42.0);
    let clamped: Vec<f32> = natural.iter().map(|&w| w.max(min_w)).collect();
    let total: f32 = clamped.iter().sum();
    if total <= avail_w {
        let scale = avail_w / total.max(1.0);
        clamped.into_iter().map(|w| w * scale).collect()
    } else {
        // Reserve min_w for every column, then distribute remaining width by excess need.
        let base_total = min_w * ncols as f32;
        let extra_budget = (avail_w - base_total).max(0.0);
        let excesses: Vec<f32> = clamped.iter().map(|&w| (w - min_w).max(0.0)).collect();
        let excess_sum: f32 = excesses.iter().sum();
        if excess_sum <= 0.0 {
            vec![avail_w / ncols as f32; ncols]
        } else {
            excesses
                .into_iter()
                .map(|ex| min_w + extra_budget * (ex / excess_sum))
                .collect()
        }
    }
}

#[allow(clippy::too_many_arguments)]
fn draw_pdf_table_row(
    ops: &mut Vec<Op>,
    wrapped_row: &[Vec<PdfLine>],
    col_widths: &[f32],
    align: &[Align],
    x: f32,
    y_top: f32,
    row_h: f32,
    is_header: bool,
    is_first_on_page: bool,
    is_striped: bool,
    size: f32,
    line_h: f32,
    pad_x: f32,
    pad_y: f32,
    palette: &PdfPalette,
    m: &Fonts,
) {
    let total_w: f32 = col_widths.iter().sum();
    let y_bottom = y_top - row_h;

    if is_header {
        fill_rect(ops, x, y_bottom, total_w, row_h, &palette.table_header_bg);
    } else if is_striped {
        fill_rect(ops, x, y_bottom, total_w, row_h, &palette.table_stripe_bg);
    }

    if is_first_on_page {
        draw_line(ops, x, y_top, x + total_w, y_top, 0.5, &palette.border);
    }
    let bottom_thickness = if is_header { 0.9 } else { 0.5 };
    draw_line(
        ops,
        x,
        y_bottom,
        x + total_w,
        y_bottom,
        bottom_thickness,
        &palette.border,
    );

    let mut cell_x = x;
    draw_line(ops, cell_x, y_bottom, cell_x, y_top, 0.5, &palette.border);
    for (c_idx, &col_w) in col_widths.iter().enumerate() {
        let cell_align = align.get(c_idx).copied().unwrap_or(Align::Left);
        if let Some(lines) = wrapped_row.get(c_idx) {
            let mut line_y_top = y_top - pad_y;
            for line in lines {
                let lw = line_width(line, size, m);
                let text_x = match cell_align {
                    Align::Left => cell_x + pad_x,
                    Align::Center => cell_x + ((col_w - lw) * 0.5).max(pad_x),
                    Align::Right => (cell_x + col_w - pad_x - lw).max(cell_x + pad_x),
                };
                let y_baseline = line_y_top - size;
                draw_styled_line(ops, line, text_x, y_baseline, size, true, palette, m);
                line_y_top -= line_h;
            }
        }
        cell_x += col_w;
        draw_line(ops, cell_x, y_bottom, cell_x, y_top, 0.5, &palette.border);
    }
}

fn write_pdf(
    blocks: &[Block],
    family: &str,
    code_family: &str,
    body: f32,
    path: &Path,
) -> Result<()> {
    let mut doc = PdfDocument::new("Notey Export");
    let m = Fonts::load(&mut doc, family, code_family)?;
    let margin = Mm(MARGIN_MM).into_pt().0;
    let page_w = Mm(PAGE_W_MM).into_pt().0;
    let page_h = Mm(PAGE_H_MM).into_pt().0;
    let palette = PdfPalette::standard();

    let mut pages: Vec<Vec<Op>> = vec![vec![]];
    let mut y = page_h - margin;
    let mut prev = None;

    for block in blocks {
        let size = block_size(block.kind, body);
        let line_h = size * line_height(block.kind);
        let base_x = margin + block.indent as f32 * INDENT_PT;
        let x = base_x + block.quote_depth as f32 * QUOTE_INDENT_PT;
        let avail_w = (page_w - margin - x).max(60.0);
        y -= gap_before(prev, block, body);
        prev = Some(block);

        if block.kind == Kind::Rule {
            let rule_h = body * 0.6;
            if y - rule_h < margin {
                pages.push(vec![]);
                y = page_h - margin;
            }
            y -= rule_h;
            let ops = pages.last_mut().unwrap();
            draw_line(ops, x, y, page_w - margin, y, 0.75, &palette.border);
            continue;
        }

        if block.kind == Kind::Table {
            let Some(table) = &block.table else {
                continue;
            };
            let ncols = table
                .align
                .len()
                .max(table.rows.iter().map(Vec::len).max().unwrap_or(0));
            if ncols == 0 || table.rows.is_empty() {
                continue;
            }
            let table_size = body * 0.92;
            let table_line_h = table_size * 1.35;
            let pad_x = 6.0;
            let pad_y = 4.5;
            let col_widths = compute_table_col_widths(table, ncols, table_size, pad_x, avail_w, &m);

            let wrapped_rows: Vec<(Vec<Vec<PdfLine>>, f32)> = table
                .rows
                .iter()
                .enumerate()
                .map(|(r_idx, row)| {
                    let force_bold = r_idx == 0;
                    let mut cells = Vec::with_capacity(ncols);
                    let mut max_lines = 1usize;
                    for (c_idx, &col_w) in col_widths.iter().enumerate() {
                        let inner_w = (col_w - pad_x * 2.0).max(12.0);
                        let lines = row.get(c_idx).map_or_else(
                            || vec![vec![]],
                            |spans| wrap_spans(spans, false, force_bold, table_size, inner_w, &m),
                        );
                        max_lines = max_lines.max(lines.len());
                        cells.push(lines);
                    }
                    let row_h = max_lines as f32 * table_line_h + pad_y * 2.0;
                    (cells, row_h)
                })
                .collect();

            let mut first_on_page = true;
            for (r_idx, (wrapped_row, row_h)) in wrapped_rows.iter().enumerate() {
                let row_h = *row_h;
                if y - row_h < margin {
                    pages.push(vec![]);
                    y = page_h - margin;
                    first_on_page = true;
                    if r_idx > 0 {
                        let (hdr_row, hdr_h) = &wrapped_rows[0];
                        let ops = pages.last_mut().unwrap();
                        if block.quote_depth > 0 {
                            draw_quote_bars(
                                ops,
                                base_x,
                                block.quote_depth,
                                y - *hdr_h,
                                *hdr_h,
                                &palette.quote_bar,
                            );
                        }
                        draw_pdf_table_row(
                            ops,
                            hdr_row,
                            &col_widths,
                            &table.align,
                            x,
                            y,
                            *hdr_h,
                            true,
                            true,
                            false,
                            table_size,
                            table_line_h,
                            pad_x,
                            pad_y,
                            &palette,
                            &m,
                        );
                        y -= *hdr_h;
                        first_on_page = false;
                    }
                }
                let ops = pages.last_mut().unwrap();
                if block.quote_depth > 0 {
                    draw_quote_bars(
                        ops,
                        base_x,
                        block.quote_depth,
                        y - row_h,
                        row_h,
                        &palette.quote_bar,
                    );
                }
                draw_pdf_table_row(
                    ops,
                    wrapped_row,
                    &col_widths,
                    &table.align,
                    x,
                    y,
                    row_h,
                    r_idx == 0,
                    first_on_page,
                    r_idx > 0 && r_idx % 2 == 0,
                    table_size,
                    table_line_h,
                    pad_x,
                    pad_y,
                    &palette,
                    &m,
                );
                y -= row_h;
                first_on_page = false;
            }
            continue;
        }

        if block.kind == Kind::Code {
            let pad_x = 7.0;
            let pad_y = 5.5;
            let lines = wrap_block(block, size, (avail_w - pad_x * 2.0).max(40.0), &m);
            let mut idx = 0;
            while idx < lines.len() {
                if y - (line_h + pad_y * 2.0) < margin {
                    pages.push(vec![]);
                    y = page_h - margin;
                }
                let fit = ((y - margin - pad_y * 2.0) / line_h).floor().max(1.0) as usize;
                let end = (idx + fit).min(lines.len());
                let slice = &lines[idx..end];
                let box_h = slice.len() as f32 * line_h + pad_y * 2.0;
                let y_bottom = y - box_h;
                let ops = pages.last_mut().unwrap();
                if block.quote_depth > 0 {
                    draw_quote_bars(
                        ops,
                        base_x,
                        block.quote_depth,
                        y_bottom,
                        box_h,
                        &palette.quote_bar,
                    );
                }
                fill_rect(ops, x, y_bottom, avail_w, box_h, &palette.code_bg);
                stroke_rect(ops, x, y_bottom, avail_w, box_h, 0.5, &palette.border);
                let mut cur_y = y - pad_y;
                for line in slice {
                    cur_y -= line_h;
                    draw_styled_line(
                        ops,
                        line,
                        x + pad_x,
                        cur_y + line_h - size,
                        size,
                        false,
                        &palette,
                        &m,
                    );
                }
                y = y_bottom;
                idx = end;
                if idx < lines.len() {
                    pages.push(vec![]);
                    y = page_h - margin;
                }
            }
            continue;
        }

        let checkbox_offset = if block.checkbox.is_some() {
            size * 0.72 + 5.0
        } else {
            0.0
        };
        let text_x = x + checkbox_offset;
        let wrap_w = (avail_w - checkbox_offset).max(40.0);
        let lines = wrap_block(block, size, wrap_w, &m);
        for (line_idx, line) in lines.iter().enumerate() {
            if y - line_h < margin {
                pages.push(vec![]);
                y = page_h - margin;
            }
            y -= line_h;
            let ops = pages.last_mut().unwrap();
            if block.quote_depth > 0 {
                draw_quote_bars(
                    ops,
                    base_x,
                    block.quote_depth,
                    y,
                    line_h,
                    &palette.quote_bar,
                );
            }
            let y_baseline = y + line_h - size;
            if line_idx == 0
                && let Some(checked) = block.checkbox
            {
                draw_checkbox(
                    ops,
                    x,
                    y_baseline,
                    size,
                    checked,
                    &palette.muted,
                    &palette.link,
                );
            }
            draw_styled_line(
                ops,
                line,
                text_x,
                y_baseline,
                size,
                block.kind != Kind::PlainCode,
                &palette,
                &m,
            );
        }
        if matches!(block.kind, Kind::Heading(1 | 2)) {
            let ops = pages.last_mut().unwrap();
            draw_line(
                ops,
                x,
                y - 4.0,
                page_w - margin,
                y - 4.0,
                0.5,
                &palette.border,
            );
            y -= 5.0;
        }
    }

    let pages = pages
        .into_iter()
        .map(|ops| PdfPage::new(Mm(PAGE_W_MM), Mm(PAGE_H_MM), ops))
        .collect();
    let mut warnings = Vec::new();
    let bytes = doc
        .with_pages(pages)
        .save(&PdfSaveOptions::default(), &mut warnings);
    std::fs::write(path, bytes).context("writing PDF file")
}

// ---- DOCX ----

fn docx_font(name: &str) -> RunFonts {
    RunFonts::new().ascii(name).hi_ansi(name).cs(name)
}

fn build_docx_run(
    span: &Span,
    family: &str,
    code_family: &str,
    size_pt: f32,
    force_bold: bool,
) -> Run {
    let half_pts = (size_pt * 2.0).round() as usize;
    let mut run = if span.text == "\n" {
        Run::new().add_break(BreakType::TextWrapping)
    } else {
        Run::new().add_text(&span.text)
    };
    run = run.size(half_pts).fonts(docx_font(if span.style.code {
        code_family
    } else {
        family
    }));
    if force_bold || span.style.bold {
        run = run.bold();
    }
    if span.style.italic {
        run = run.italic();
    }
    if span.style.strike {
        run = run.strike();
    }
    if span.style.link.is_some() {
        run = run.color("0969DA").underline("single");
    } else if span.style.muted {
        run = run.color("616B7A");
    }
    if span.style.code {
        run = run.shading(Shading::new().shd_type(ShdType::Clear).fill("EEF0F4"));
    }
    run
}

fn add_spans_to_paragraph(
    mut p: Paragraph,
    spans: &[Span],
    family: &str,
    code_family: &str,
    size_pt: f32,
    force_bold: bool,
) -> Paragraph {
    for span in spans {
        let run = build_docx_run(span, family, code_family, size_pt, force_bold);
        if let Some(url) = &span.style.link
            && !url.is_empty()
            && span.text != "\n"
        {
            p = p.add_hyperlink(Hyperlink::new(url, HyperlinkType::External).add_run(run));
        } else {
            p = p.add_run(run);
        }
    }
    p
}

/// Word gets the font by name; it substitutes when the reader doesn't have it installed.
fn write_docx(
    blocks: &[Block],
    family: &str,
    code_family: &str,
    body: f32,
    path: &Path,
) -> Result<()> {
    let file = std::fs::File::create(path).context("creating DOCX file")?;
    let half_pts = |pt: f32| (pt * 2.0).round() as usize;
    // Word's units: spacing in twentieths of a point, line height in 240ths of a line.
    // Word's line is the font's height (about 1.2em), hence the division.
    let spacing = |kind: Kind, before: f32, after: f32| {
        LineSpacing::new()
            .before((before * 20.0).round() as u32)
            .after((after * 20.0).round() as u32)
            .line((line_height(kind) / 1.2 * 240.0).round() as i32)
            .line_rule(LineSpacingType::Auto)
    };
    let mut docx = Docx::new();
    let mut prev = None;
    for (i, block) in blocks.iter().enumerate() {
        let gap = gap_before(prev, block, body);
        prev = Some(block);
        // Word tables can't have space before them, so it goes after the previous block.
        let after = match blocks.get(i + 1) {
            Some(next) if next.kind == Kind::Table => gap_before(prev, next, body),
            _ => 0.0,
        };
        let total_indent = (block.indent as i32 + block.quote_depth as i32) * 360;

        if block.kind == Kind::Table {
            let Some(table) = &block.table else {
                continue;
            };
            let ncols = table
                .align
                .len()
                .max(table.rows.iter().map(Vec::len).max().unwrap_or(0));
            if ncols == 0 || table.rows.is_empty() {
                continue;
            }
            let table_size = body * 0.92;
            let col_w_dxa = (9000 / ncols).max(1000);
            let mut docx_rows = Vec::with_capacity(table.rows.len());
            for (r_idx, row) in table.rows.iter().enumerate() {
                let is_header = r_idx == 0;
                let mut docx_cells = Vec::with_capacity(ncols);
                for c_idx in 0..ncols {
                    let align = match table.align.get(c_idx).copied().unwrap_or(Align::Left) {
                        Align::Left => AlignmentType::Left,
                        Align::Center => AlignmentType::Center,
                        Align::Right => AlignmentType::Right,
                    };
                    let mut p = Paragraph::new()
                        .align(align)
                        .line_spacing(spacing(block.kind, 0.0, 0.0));
                    if let Some(spans) = row.get(c_idx) {
                        p = add_spans_to_paragraph(
                            p,
                            spans,
                            family,
                            code_family,
                            table_size,
                            is_header,
                        );
                    }
                    let mut cell = TableCell::new()
                        .width(col_w_dxa, WidthType::Dxa)
                        .add_paragraph(p);
                    if is_header {
                        cell = cell.shading(Shading::new().shd_type(ShdType::Clear).fill("EEF1F5"));
                    } else if r_idx % 2 == 0 {
                        cell = cell.shading(Shading::new().shd_type(ShdType::Clear).fill("FAFBFC"));
                    }
                    docx_cells.push(cell);
                }
                docx_rows.push(TableRow::new(docx_cells));
            }
            let border = TableBorder::new(docx_rs::TableBorderPosition::Top)
                .border_type(BorderType::Single)
                .size(4)
                .color("CCD3DB");
            let borders = TableBorders::new()
                .set(border)
                .set(
                    TableBorder::new(docx_rs::TableBorderPosition::Bottom)
                        .border_type(BorderType::Single)
                        .size(4)
                        .color("CCD3DB"),
                )
                .set(
                    TableBorder::new(docx_rs::TableBorderPosition::Left)
                        .border_type(BorderType::Single)
                        .size(4)
                        .color("CCD3DB"),
                )
                .set(
                    TableBorder::new(docx_rs::TableBorderPosition::Right)
                        .border_type(BorderType::Single)
                        .size(4)
                        .color("CCD3DB"),
                )
                .set(
                    TableBorder::new(docx_rs::TableBorderPosition::InsideH)
                        .border_type(BorderType::Single)
                        .size(4)
                        .color("CCD3DB"),
                )
                .set(
                    TableBorder::new(docx_rs::TableBorderPosition::InsideV)
                        .border_type(BorderType::Single)
                        .size(4)
                        .color("CCD3DB"),
                );
            let mut docx_table = Table::new(docx_rows)
                .set_borders(borders)
                .margins(docx_rs::TableCellMargins::new().margin(80, 120, 80, 120));
            if total_indent > 0 {
                docx_table = docx_table.indent(total_indent);
            }
            docx = docx.add_table(docx_table);
            continue;
        }

        // One paragraph per code line keeps Word's spacing from double-spacing code.
        if matches!(block.kind, Kind::Code | Kind::PlainCode) {
            let text = &block.spans[0].text;
            let code_size = block_size(block.kind, body);
            let is_fenced = block.kind == Kind::Code;
            let mut lines = text
                .strip_suffix('\n')
                .unwrap_or(text)
                .split('\n')
                .enumerate()
                .peekable();
            while let Some((i, line)) = lines.next() {
                let before = if i == 0 { gap } else { 0.0 };
                let after = if lines.peek().is_none() { after } else { 0.0 };
                let mut run = Run::new()
                    .add_text(line)
                    .fonts(docx_font(code_family))
                    .size(half_pts(code_size));
                if is_fenced {
                    run = run.shading(Shading::new().shd_type(ShdType::Clear).fill("F3F5F7"));
                }
                let p = Paragraph::new()
                    .line_spacing(spacing(block.kind, before, after))
                    .add_run(run)
                    .indent(Some(total_indent), None, None, None);
                docx = docx.add_paragraph(p);
            }
            continue;
        }

        let mut p = Paragraph::new()
            .line_spacing(spacing(block.kind, gap, after))
            .indent(Some(total_indent), None, None, None);
        if block.kind == Kind::Rule {
            p = p.add_run(
                Run::new()
                    .add_text("\u{2500}".repeat(36))
                    .color("CCD3DB")
                    .fonts(docx_font(family)),
            );
            docx = docx.add_paragraph(p);
            continue;
        }
        if block.quote_depth > 0 {
            p = p.add_run(
                Run::new()
                    .add_text("\u{2502} ".repeat(block.quote_depth as usize))
                    .color("BFC7D1")
                    .fonts(docx_font(family))
                    .size(half_pts(block_size(block.kind, body))),
            );
        }
        if let Some(checked) = block.checkbox {
            p = p.add_run(
                Run::new()
                    .add_text(if checked { "\u{2611} " } else { "\u{2610} " })
                    .color(if checked { "0969DA" } else { "616B7A" })
                    .fonts(docx_font(family))
                    .size(half_pts(block_size(block.kind, body))),
            );
        }
        let heading = matches!(block.kind, Kind::Heading(_));
        p = add_spans_to_paragraph(
            p,
            &block.spans,
            family,
            code_family,
            block_size(block.kind, body),
            heading,
        );
        docx = docx.add_paragraph(p);
    }
    docx.build()
        .pack(file)
        .map_err(|err| anyhow::anyhow!("{err:?}"))
        .context("writing DOCX file")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn markdown_becomes_styled_blocks() {
        let blocks = markdown_blocks(
            "# Title\n\nSome **bold**, *italic*, ~~strike~~, `code`, and [link](https://example.com).\n\n- one\n- [x] done\n- [ ] todo\n\n> Quoted\n\n| Name | Score |\n| :--- | ---: |\n| Alice | **100** |\n\n---\n",
        );
        assert_eq!(blocks[0].kind, Kind::Heading(1));
        assert_eq!(blocks[0].spans[0].text, "Title");
        assert!(
            blocks[1]
                .spans
                .iter()
                .any(|s| s.text == "bold" && s.style.bold)
        );
        assert!(
            blocks[1]
                .spans
                .iter()
                .any(|s| s.text == "italic" && s.style.italic)
        );
        assert!(
            blocks[1]
                .spans
                .iter()
                .any(|s| s.text == "strike" && s.style.strike)
        );
        assert!(
            blocks[1]
                .spans
                .iter()
                .any(|s| s.text == "code" && s.style.code)
        );
        assert!(blocks[1].spans.iter().any(|s| {
            s.text == "link" && s.style.link.as_deref() == Some("https://example.com")
        }));
        assert_eq!(blocks[2].spans[0].text, "\u{2022} ");
        assert_eq!(blocks[2].indent, 1);
        assert_eq!(blocks[3].checkbox, Some(true));
        assert_eq!(blocks[4].checkbox, Some(false));
        assert_eq!(blocks[5].quote_depth, 1);
        assert_eq!(blocks[6].kind, Kind::Table);
        let table = blocks[6].table.as_ref().unwrap();
        assert_eq!(table.align, vec![Align::Left, Align::Right]);
        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.rows[0][0][0].text, "Name");
        assert!(table.rows[1][1][0].style.bold);
        assert_eq!(blocks[7].kind, Kind::Rule);
        // No raw markdown syntax leaks through.
        assert!(
            !blocks
                .iter()
                .flat_map(|b| &b.spans)
                .any(|s| s.text.contains(['#', '*', '`', '~']))
        );
    }

    #[test]
    fn resolves_reference_links_and_footnotes() {
        let blocks = markdown_blocks(
            "See [docs][d] and note[^1].\n\n[d]: https://example.org\n[^1]: Footnote body.\n",
        );
        assert!(blocks[0].spans.iter().any(|s| {
            s.text == "docs" && s.style.link.as_deref() == Some("https://example.org")
        }));
        assert!(blocks[0].spans.iter().any(|s| s.text == "[1]"));
        assert!(blocks.iter().any(|b| b.kind == Kind::Rule));
        assert!(
            blocks
                .iter()
                .any(|b| b.spans.iter().any(|s| s.text == "Footnote body."))
        );
    }

    #[test]
    fn heading_scale_and_block_gaps() {
        assert!(HEADING_SCALE.is_sorted_by(|a, b| a > b));
        let blocks = markdown_blocks("# A\n\ntext\n\n## B\n\ntext\n\n- one\n- two\n");
        let gap = |i: usize| gap_before(i.checked_sub(1).map(|p| &blocks[p]), &blocks[i], 12.0);
        assert_eq!(gap(0), 0.0);
        assert!(gap(2) > gap(3));
        assert!(gap(5) < gap(4));
    }

    #[test]
    fn wraps_within_width() {
        let m = Fonts::load(&mut PdfDocument::new("t"), "Lilex", "Lilex").unwrap();
        let block = markdown_blocks(&("word ".repeat(200) + &"x".repeat(300))).remove(0);
        let lines = wrap_block(&block, 11.0, 400.0, &m);
        assert!(lines.len() > 5);
        for line in &lines {
            let w = line_width(line, 11.0, &m);
            assert!(w <= 400.0, "line too wide: {w}");
        }
    }

    #[test]
    fn finds_system_font_faces() {
        // Menlo ships as a .ttc; its bold must be a different face than regular.
        let [r, b, ..] = family_faces("Menlo");
        assert!(matches!(r, Source::File(..)));
        assert_ne!(r, b);
        assert_eq!(
            family_faces("No Such Font")[0],
            Source::Lilex { bold: false }
        );
    }

    #[test]
    fn exports_pdf_and_docx_all_markdown_components() {
        let dir = std::env::temp_dir();
        let text = [
            "# Heading 1",
            "## Heading 2",
            "### Heading 3",
            "#### Heading 4",
            "##### Heading 5",
            "###### Heading 6",
            "",
            "Paragraph with **bold**, *italic*, ***both***, ~~strike~~, `inline code`, [link](https://example.com), [ref][r], ![diagram](img.png), and footnote[^1].",
            "",
            "Line one<br>Line two",
            "",
            "> Blockquote line",
            "> > Nested quote with **bold**",
            "",
            "- Bullet 1",
            "  - Nested bullet",
            "1. Ordered 1",
            "2. Ordered 2",
            "- [x] Completed task",
            "- [ ] Pending task",
            "",
            "```rust",
            "fn main() {",
            "    println!(\"hello world\");",
            "}",
            "```",
            "",
            "| Feature | Alignment | Status | Notes |",
            "| :--- | :---: | ---: | :--- |",
            "| Tables | Center | 100% | Supports **bold**, `code`, and long wrapping text across columns |",
            "| Quotes | Center | 100% | Nested bars and muted text |",
            "| Links | Center | 100% | [Example](https://example.com) |",
            "",
            "---",
            "",
            "[r]: https://rust-lang.org",
            "[^1]: Footnote explanation text.",
        ]
        .join("\n");

        for (markdown, family) in [
            (false, "Menlo"),
            (true, "Helvetica Neue"),
            (true, "Lilex"),
            (false, "No Such Font"),
        ] {
            let write =
                |format: ExportFormat, path| format.write(&text, markdown, family, 12.0, path);
            let pdf_path = dir.join("notey_export_test.pdf");
            write(ExportFormat::Pdf, &pdf_path).unwrap();
            assert!(std::fs::read(&pdf_path).unwrap().starts_with(b"%PDF"));
            std::fs::remove_file(&pdf_path).unwrap();

            let docx_path = dir.join("notey_export_test.docx");
            write(ExportFormat::Docx, &docx_path).unwrap();
            assert!(std::fs::read(&docx_path).unwrap().starts_with(b"PK")); // docx is a zip archive
            std::fs::remove_file(&docx_path).unwrap();
        }
    }
}
