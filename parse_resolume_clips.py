#!/usr/bin/env python3
"""
Resolume Arena – экспорт клипов в Excel
"""

import os
import sys
import subprocess
from xml.etree import ElementTree as ET
from xml.etree.ElementTree import Element
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Any

# ─── Цветовая схема (тёмная) ─────────────────────────────────────────────────
BG        = '#0f1117'   # фон окна
SURFACE   = '#1a1d27'   # поверхность карточек
BORDER    = '#2a2d3a'   # граница
ACCENT    = '#6c63ff'   # фиолетовый акцент
ACCENT_HV = '#8b84ff'   # hover
TEXT      = '#e8eaf0'   # основной текст
SUBTEXT   = '#8b8fa8'   # вторичный текст
SUCCESS   = '#4ade80'   # зелёный
DANGER    = '#f87171'   # красный
FONT_UI   = 'Segoe UI'


# ─── XML helpers ─────────────────────────────────────────────────────────────

def get_clip_name(clip: Element) -> str:
    p = clip.find('Params[@name="Params"]/Param[@name="Name"]')
    if p is not None:
        return p.get('value') or p.get('default') or ''
    return clip.get('name', '')


def get_clip_duration(clip: Element) -> float | None:
    try:
        pos = clip.find('Transport/Params/ParamRange[@name="Position"]')
        if pos is None:
            return None
        tl = pos.find('PhaseSourceTransportTimeline')
        if tl is None:
            return None
        dur = tl.find('Params/ParamRange[@name="Duration"]')
        if dur is None:
            return None
        val = dur.get('value') or dur.get('default')
        if val:
            return round(float(val), 3)
    except Exception:
        pass
    return None


def get_autopilot(clip: Element) -> bool:
    """
    Autopilot активен когда Target != default (т.е. указан конкретный целевой клип).
    Пустой/нулевой Target = AAAAAAAAAAA=,AAAAAAAAAAA= означает автопилот выключен.
    """
    try:
        target = clip.find('Params[@name="AutoPilot"]/Param[@name="Target"]')
        if target is not None:
            val     = target.get('value', '')
            default = target.get('default', '')
            # Если value совпадает с default — автопилот не настроен
            if val and val != default:
                return True
    except Exception:
        pass
    return False


def parse_clips(xml_text: str) -> list[dict[str, Any]]:
    root = ET.fromstring(xml_text)
    rows = []
    for clip in root.findall('.//Clip'):
        rows.append({
            'name':      get_clip_name(clip),
            'layer':     clip.get('layerIndex', ''),
            'column':    clip.get('columnIndex', ''),
            'duration':  get_clip_duration(clip),
            'autopilot': get_autopilot(clip),
        })
    return rows


# ─── Статистика ──────────────────────────────────────────────────────────────

def fmt_hms(sec: float | int) -> str:
    """Секунды → ЧЧ:ММ:СС"""
    sec = int(sec)
    h = sec // 3600
    m = (sec % 3600) // 60
    s = sec % 60
    return f'{h:02d}:{m:02d}:{s:02d}'


def compute_stats(rows: list[dict[str, Any]]) -> dict[str, Any]:
    """Вычисляет сводную статистику по клипам."""
    from collections import defaultdict

    total_sec = round(sum(r['duration'] for r in rows if r['duration'] is not None), 3)
    autopilot_count = sum(1 for r in rows if r['autopilot'])

    layer_dur   = defaultdict(float)
    layer_count = defaultdict(int)
    for r in rows:
        lyr = r['layer']
        layer_count[lyr] += 1
        if r['duration'] is not None:
            layer_dur[lyr] += r['duration']

    layers = sorted(layer_count.keys(),
                    key=lambda x: int(x) if str(x).isdigit() else x)

    return {
        'total_sec':       total_sec,
        'total_hms':       fmt_hms(total_sec),
        'autopilot_count': autopilot_count,
        'clip_count':      len(rows),
        'layers':          layers,
        'layer_dur':       {k: round(v, 3) for k, v in layer_dur.items()},
        'layer_count':     dict(layer_count),
    }


# ─── Export ──────────────────────────────────────────────────────────────────

def export_xlsx(rows: list[dict[str, Any]], out_path: str) -> None:
    import xlsxwriter  # type: ignore[import-untyped]

    wb = xlsxwriter.Workbook(out_path)
    ws = wb.add_worksheet('Clips')

    # ── Форматы ─────────────────────────────────────────────────────────────
    hdr_fmt = wb.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 10,
        'font_color': '#FFFFFF', 'bg_color': '#1a1d27',
        'align': 'center', 'valign': 'vcenter',
        'border': 0,
    })
    cell_fmt = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 10,
        'valign': 'vcenter', 'align': 'left',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })
    cell_fmt_c = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 10,
        'valign': 'vcenter', 'align': 'center',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })
    dur_fmt = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 10,
        'valign': 'vcenter', 'align': 'center',
        'num_format': '0.000',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })
    alt_fmt = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 10,
        'valign': 'vcenter', 'align': 'left',
        'bg_color': '#F5F6FA',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })
    alt_fmt_c = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 10,
        'valign': 'vcenter', 'align': 'center',
        'bg_color': '#F5F6FA',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })
    alt_dur_fmt = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 10,
        'valign': 'vcenter', 'align': 'center',
        'bg_color': '#F5F6FA',
        'num_format': '0.000',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })

    # ── Заголовки ───────────────────────────────────────────────────────────
    headers = ['Имя', 'Слой', 'Колонка', 'Duration (сек)', 'Автопилот']
    ws.set_row(0, 22, hdr_fmt)
    for col, h in enumerate(headers):
        ws.write(0, col, h, hdr_fmt)

    # ── Данные ──────────────────────────────────────────────────────────────
    col_widths = [len(h) for h in headers]
    for i, r in enumerate(rows):
        row = i + 1
        alt = (row % 2 == 0)
        f  = alt_fmt   if alt else cell_fmt
        fc = alt_fmt_c if alt else cell_fmt_c
        fd = alt_dur_fmt if alt else dur_fmt

        ws.set_row(row, 18)
        ws.write(row, 0, r['name'], f)
        layer_val  = int(r['layer'])  + 1 if str(r['layer']).isdigit()  else r['layer']
        col_val    = int(r['column']) + 1 if str(r['column']).isdigit() else r['column']
        ws.write(row, 1, layer_val,  fc)
        ws.write(row, 2, col_val,    fc)
        if r['duration'] is not None:
            ws.write_number(row, 3, r['duration'], fd)
        else:
            ws.write(row, 3, '', fc)
        ws.write(row, 4, 'Вкл' if r['autopilot'] else 'Выкл', fc)

        col_widths[0] = max(col_widths[0], len(str(r['name'])))
        col_widths[3] = max(col_widths[3], 10)

    # ── Ширина колонок А–Е ──────────────────────────────────────────────────
    widths = [col_widths[0] + 4, 8, 10, 16, 10]
    for col, w in enumerate(widths):
        ws.set_column(col, col, w)

    # ── Колонка F — пустой разделитель ──────────────────────────────────────
    ws.set_column(5, 5, 3)

    # ── Сайдбар статистики: G (6) = название, H (7) = значение/формула ──────
    stats = compute_stats(rows)

    # Форматы сайдбара
    sb_hdr = wb.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 10,
        'font_color': '#FFFFFF', 'bg_color': '#1a1d27',
        'align': 'left', 'valign': 'vcenter', 'border': 0,
    })
    sb_sec_hdr = wb.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 9,
        'font_color': '#6c63ff', 'bg_color': '#F5F6FA',
        'top': 1, 'top_color': '#6c63ff', 'valign': 'vcenter',
    })
    sb_lbl = wb.add_format({
        'font_name': 'Segoe UI', 'font_size': 10,
        'font_color': '#8b8fa8', 'valign': 'vcenter',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })
    sb_val = wb.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 10,
        'valign': 'vcenter', 'bottom': 1, 'bottom_color': '#E0E4F0',
    })
    sb_val3 = wb.add_format({
        'bold': True, 'font_name': 'Segoe UI', 'font_size': 10,
        'num_format': '0.000', 'valign': 'vcenter',
        'bottom': 1, 'bottom_color': '#E0E4F0',
    })

    ws.set_column(6, 6, 34)   # G — метки
    ws.set_column(7, 7, 18)   # H — значения

    # Формула ЧЧ:ММ:СС из ячейки (напр. "H4") с секундами
    def hms_f(ref: str) -> str:
        return (
            f'=TEXT(INT({ref}/3600),"00")&":"'
            f'&TEXT(INT(MOD({ref},3600)/60),"00")&":"'
            f'&TEXT(INT(MOD({ref},60)),"00")'
        )

    def sb_row(r: int, label: str, formula: str, num: bool = False) -> None:
        ws.set_row(r, 20)
        ws.write(r, 6, label, sb_lbl)
        ws.write_formula(r, 7, formula, sb_val3 if num else sb_val)

    r = 0
    # ── ОБЩАЯ СТАТИСТИКА ────────────────────────────────────────────────────
    ws.set_row(r, 22)
    ws.write(r, 6, 'ОБЩАЯ СТАТИСТИКА', sb_hdr)
    ws.write(r, 7, '', sb_hdr); r += 1

    sb_row(r, 'Всего клипов',          '=COUNTA(A:A)-1'); r += 1
    sb_row(r, 'Автопилот включён',     '=COUNTIF(E:E,"Вкл")'); r += 1
    sb_row(r, 'Общее время (секунды)', '=SUM(D:D)', num=True)
    total_ref = f'H{r + 1}'; r += 1
    sb_row(r, 'Общее время (ЧЧ:ММ:СС)', hms_f(total_ref)); r += 1

    r += 1  # пустая строка-разделитель

    # ── ПО СЛОЯМ ────────────────────────────────────────────────────────────
    ws.set_row(r, 22)
    ws.write(r, 6, 'ПО СЛОЯМ', sb_hdr)
    ws.write(r, 7, '', sb_hdr); r += 1

    for lyr in stats['layers']:
        lyr_ui = int(lyr) + 1 if str(lyr).isdigit() else lyr

        ws.set_row(r, 22)
        ws.write(r, 6, f'Слой {lyr_ui}', sb_sec_hdr)
        ws.write(r, 7, '', sb_sec_hdr); r += 1

        sb_row(r, f'  Клипов в слое {lyr_ui}',
               f'=COUNTIF(B:B,{lyr_ui})'); r += 1

        sb_row(r, f'  Время слоя {lyr_ui} (секунды)',
               f'=SUMIF(B:B,{lyr_ui},D:D)', num=True)
        layer_ref = f'H{r + 1}'; r += 1

        sb_row(r, f'  Время слоя {lyr_ui} (ЧЧ:ММ:СС)',
               hms_f(layer_ref)); r += 1

    wb.close()


def open_file_with_default_app(path: str) -> None:
    try:
        if os.name == 'nt':
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.run(['open', path])
        else:
            subprocess.run(['xdg-open', path])
    except Exception:
        pass


# ─── Обработка ───────────────────────────────────────────────────────────────

def process(xml_text: str, source_path: str | None, status_cb: Any, stats_cb: Any = None) -> None:
    status_cb('Разбор XML...', SUBTEXT)
    try:
        rows = parse_clips(xml_text)
    except ET.ParseError as e:
        status_cb(f'Ошибка XML: {e}', DANGER)
        messagebox.showerror('Ошибка XML', f'Не удалось разобрать XML:\n{e}')
        return

    if not rows:
        status_cb('Клипы не найдены.', DANGER)
        messagebox.showwarning('Нет данных', 'Клипы не найдены в файле.')
        return

    stats = compute_stats(rows)
    if stats_cb:
        stats_cb(stats)

    # Сохраняем рядом с EXE (при запуске из PyInstaller) или рядом со скриптом
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(exe_dir, 'clips_output.xlsx')

    status_cb(f'Экспорт {len(rows)} клипов...', SUBTEXT)

    # Если файл занят — пробуем сохранить под именем с номером
    for attempt in range(10):
        candidate = out_path if attempt == 0 else out_path.replace('.xlsx', f'_{attempt}.xlsx')
        try:
            export_xlsx(rows, candidate)
            out_path = candidate
            break
        except PermissionError:
            if attempt == 0:
                # Сначала предлагаем закрыть файл
                retry = messagebox.askretrycancel(
                    'Файл занят',
                    f'Не удалось сохранить файл — он открыт в другой программе.\n\n'
                    f'{candidate}\n\n'
                    f'Закройте файл и нажмите «Повторить», или «Отмена» для сохранения копии.'
                )
                if retry:
                    try:
                        export_xlsx(rows, candidate)
                        out_path = candidate
                        break
                    except PermissionError:
                        pass  # попробуем с суффиксом
            continue
        except Exception as e:
            status_cb(f'Ошибка экспорта: {e}', DANGER)
            messagebox.showerror('Ошибка экспорта', f'Не удалось создать XLSX:\n{e}')
            return
    else:
        status_cb('Не удалось сохранить файл — закройте Excel и попробуйте снова.', DANGER)
        messagebox.showerror('Ошибка', 'Не удалось сохранить файл.\nЗакройте Excel и попробуйте снова.')
        return

    status_cb(f'Готово — экспортировано {len(rows)} клипов', SUCCESS)
    open_file_with_default_app(out_path)


# ─── UI helpers ──────────────────────────────────────────────────────────────

def make_button(parent: tk.Widget, text: str, command: Any, primary: bool = False, width: int | None = None) -> tk.Label:
    bg = ACCENT if primary else SURFACE
    fg = '#ffffff'
    btn = tk.Label(
        parent, text=text,
        font=(FONT_UI, 10, 'bold' if primary else 'normal'),
        bg=bg, fg=fg,
        padx=18, pady=10,
        cursor='hand2',
        relief='flat',
    )
    if width:
        btn.config(width=width)

    def on_enter(_):
        btn.config(bg=ACCENT_HV if primary else BORDER)
    def on_leave(_):
        btn.config(bg=ACCENT if primary else SURFACE)
    def on_click(_):
        command()

    btn.bind('<Enter>', on_enter)
    btn.bind('<Leave>', on_leave)
    btn.bind('<Button-1>', on_click)
    return btn


# ─── GUI ─────────────────────────────────────────────────────────────────────

def run_gui():
    win = tk.Tk()
    win.title('Resolume Clip Exporter')
    win.resizable(False, False)
    win.configure(bg=BG)
    win.geometry('560x620')

    # ── Центрируем окно ────────────────────────────────────────────────────
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    ww, wh = 560, 620
    win.geometry(f'{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}')

    state: dict[str, str | None] = {'filepath': None, 'xml': None}

    # ── Шапка ──────────────────────────────────────────────────────────────
    header = tk.Frame(win, bg=SURFACE, pady=0)
    header.pack(fill='x')
    tk.Frame(header, bg=ACCENT, width=4, height=56).pack(side='left', fill='y')
    title_box = tk.Frame(header, bg=SURFACE, padx=16, pady=14)
    title_box.pack(side='left', fill='both', expand=True)
    tk.Label(title_box, text='Resolume Clip Exporter',
             font=(FONT_UI, 14, 'bold'), bg=SURFACE, fg=TEXT).pack(anchor='w')
    tk.Label(title_box, text='Экспорт клипов Arena в Excel (XLSX)',
             font=(FONT_UI, 9), bg=SURFACE, fg=SUBTEXT).pack(anchor='w')

    # ── Тело ───────────────────────────────────────────────────────────────
    body = tk.Frame(win, bg=BG, padx=24, pady=20)
    body.pack(fill='both', expand=True)

    # Секция: выбор файла
    tk.Label(body, text='ИСТОЧНИК', font=(FONT_UI, 8, 'bold'),
             bg=BG, fg=SUBTEXT).pack(anchor='w', pady=(0, 6))

    file_row = tk.Frame(body, bg=SURFACE, padx=12, pady=10,
                        highlightbackground=BORDER, highlightthickness=1)
    file_row.pack(fill='x')

    lbl_file = tk.Label(file_row, text='Файл не выбран',
                        font=(FONT_UI, 9), bg=SURFACE, fg=SUBTEXT,
                        anchor='w', width=38)
    lbl_file.pack(side='left', fill='x', expand=True)

    def btn_open():
        path = filedialog.askopenfilename(
            title='Выберите файл проекта Resolume',
            filetypes=[('Text / XML', '*.txt;*.xml'), ('Все файлы', '*.*')]
        )
        if not path:
            return
        state['filepath'] = path
        try:
            with open(path, 'r', encoding='utf-8', errors='replace') as f:
                xml_text = f.read()
        except Exception as e:
            messagebox.showerror('Ошибка', str(e))
            return
        state['xml'] = xml_text
        name = os.path.basename(path)
        lbl_file.config(text=name, fg=SUCCESS)
        lbl_charcount.config(text=f'{len(xml_text):,} символов', fg=SUBTEXT)
        set_status('Файл загружен: ' + name, SUCCESS)

    open_btn = make_button(file_row, '  Обзор…', btn_open)
    open_btn.pack(side='right')

    # Секция: или вставить
    div = tk.Frame(body, bg=BG, pady=12)
    div.pack(fill='x')
    tk.Frame(div, bg=BORDER, height=1).pack(fill='x', side='left', expand=True, pady=7)
    tk.Label(div, text='  или вставьте XML  ', font=(FONT_UI, 8),
             bg=BG, fg=SUBTEXT).pack(side='left')
    tk.Frame(div, bg=BORDER, height=1).pack(fill='x', side='left', expand=True, pady=7)

    paste_frame = tk.Frame(body, bg=SURFACE,
                           highlightbackground=BORDER, highlightthickness=1)
    paste_frame.pack(fill='x')
    paste_area = tk.Text(paste_frame, height=5, bg=SURFACE, fg=TEXT,
                         insertbackground=TEXT,
                         font=('Consolas', 8), relief='flat',
                         padx=10, pady=8, wrap='none',
                         selectbackground=ACCENT)
    scroll_y = tk.Scrollbar(paste_frame, orient='vertical',
                            command=paste_area.yview, bg=BORDER,
                            troughcolor=SURFACE, relief='flat')
    paste_area.config(yscrollcommand=scroll_y.set)
    scroll_y.pack(side='right', fill='y')
    paste_area.pack(fill='both', expand=True)

    # Placeholder
    PLACEHOLDER = 'Вставьте XML-текст сюда (Ctrl+V)…'
    paste_area.insert('1.0', PLACEHOLDER)
    paste_area.config(fg=SUBTEXT)

    def _clear_placeholder():
        if paste_area.get('1.0', 'end-1c') == PLACEHOLDER:
            paste_area.delete('1.0', tk.END)
            paste_area.config(fg=TEXT)

    def _sync_state():
        content = paste_area.get('1.0', 'end-1c').strip()
        if content and content != PLACEHOLDER:
            state['xml'] = content
            lbl_file.config(text='Текст вставлен вручную', fg=SUBTEXT)
            lbl_charcount.config(text=f'{len(content):,} символов', fg=SUBTEXT)

    def on_focus_in(_: Any) -> None:
        _clear_placeholder()

    def on_paste(event: Any) -> str:
        _clear_placeholder()
        try:
            clip = paste_area.clipboard_get()
        except tk.TclError:
            return 'break'
        try:
            # удалить выделенный текст перед вставкой (если есть)
            paste_area.delete(tk.SEL_FIRST, tk.SEL_LAST)
        except tk.TclError:
            pass
        paste_area.insert(tk.INSERT, clip)
        _sync_state()
        return 'break'  # предотвратить двойную вставку от дефолтного обработчика

    def on_key_release(_: Any) -> None:
        _sync_state()

    paste_area.bind('<FocusIn>',    on_focus_in)
    paste_area.bind('<<Paste>>',    on_paste)
    paste_area.bind('<Control-v>',  on_paste)
    paste_area.bind('<Control-V>',  on_paste)
    paste_area.bind('<KeyRelease>', on_key_release)

    # Нижняя строка информации
    info_row = tk.Frame(body, bg=BG, pady=2)
    info_row.pack(fill='x')
    lbl_charcount = tk.Label(info_row, text='', font=(FONT_UI, 8),
                             bg=BG, fg=SUBTEXT)
    lbl_charcount.pack(side='right')

    # ── Панель статистики ──────────────────────────────────────────────────
    tk.Frame(body, bg=BORDER, height=1).pack(fill='x', pady=(6, 0))
    stats_outer = tk.Frame(body, bg=BG, pady=6)
    stats_outer.pack(fill='x')
    tk.Label(stats_outer, text='СТАТИСТИКА', font=(FONT_UI, 8, 'bold'),
             bg=BG, fg=SUBTEXT).pack(anchor='w', pady=(0, 4))

    stats_top = tk.Frame(stats_outer, bg=BG)
    stats_top.pack(fill='x')

    # StringVars для обновления без пересоздания виджетов
    sv_clips    = tk.StringVar(value='—')
    sv_ap       = tk.StringVar(value='—')
    sv_sec      = tk.StringVar(value='—')
    sv_hms      = tk.StringVar(value='—')
    sv_layers   = tk.StringVar(value='—')

    def mk_stat_col(parent, label, var):
        f = tk.Frame(parent, bg=BG)
        f.pack(side='left', padx=(0, 20))
        tk.Label(f, text=label, font=(FONT_UI, 8), bg=BG, fg=SUBTEXT).pack(anchor='w')
        tk.Label(f, textvariable=var, font=(FONT_UI, 10, 'bold'),
                 bg=BG, fg=TEXT).pack(anchor='w')

    mk_stat_col(stats_top, 'Клипов',         sv_clips)
    mk_stat_col(stats_top, 'Автопилот ВКЛ',  sv_ap)
    mk_stat_col(stats_top, 'Время (сек)',     sv_sec)
    mk_stat_col(stats_top, 'ЧЧ:ММ:СС',       sv_hms)

    lbl_layers_detail = tk.Label(stats_outer, textvariable=sv_layers,
                                  font=(FONT_UI, 8), bg=BG, fg=SUBTEXT,
                                  justify='left', anchor='w')
    lbl_layers_detail.pack(fill='x', pady=(4, 0))

    def set_stats(stats):
        sv_clips.set(str(stats['clip_count']))
        sv_ap.set(str(stats['autopilot_count']))
        sv_sec.set(f"{stats['total_sec']:.3f}")
        sv_hms.set(stats['total_hms'])
        lines = []
        for lyr in stats['layers']:
            lyr_ui = int(lyr) + 1 if str(lyr).isdigit() else lyr
            cnt    = stats['layer_count'][lyr]
            dur    = stats['layer_dur'].get(lyr, 0.0)
            lines.append(
                f'Слой {lyr_ui}:  {cnt} кл.  |  {dur:.3f} сек  |  {fmt_hms(dur)}'
            )
        sv_layers.set('\n'.join(lines))

    # ── Статус ─────────────────────────────────────────────────────────────
    status_var = tk.StringVar(value='Ожидание…')
    status_bar = tk.Frame(win, bg=SURFACE, pady=0,
                          highlightbackground=BORDER, highlightthickness=1)
    status_bar.pack(fill='x', side='bottom')
    tk.Frame(status_bar, bg=BG, width=4).pack(side='left', fill='y')
    lbl_status_dot = tk.Label(status_bar, text='●', font=(FONT_UI, 8),
                              bg=SURFACE, fg=SUBTEXT)
    lbl_status_dot.pack(side='left', padx=(6, 4), pady=8)
    lbl_status = tk.Label(status_bar, textvariable=status_var,
                          font=(FONT_UI, 9), bg=SURFACE, fg=SUBTEXT, anchor='w')
    lbl_status.pack(side='left', fill='x', expand=True, pady=8)

    def set_status(msg, color=SUBTEXT):
        status_var.set(msg)
        lbl_status.config(fg=color)
        lbl_status_dot.config(fg=color)

    # ── Кнопка Экспортировать ──────────────────────────────────────────────
    bottom = tk.Frame(win, bg=BG, padx=24, pady=12)
    bottom.pack(fill='x', side='bottom')

    def btn_process():
        xml_text = (state.get('xml') or '').strip()
        if not xml_text or xml_text == PLACEHOLDER:
            # Try paste area as fallback
            xml_text = paste_area.get('1.0', 'end-1c').strip()
        if not xml_text or xml_text == PLACEHOLDER:
            set_status('Выберите файл или вставьте XML.', DANGER)
            return
        process(xml_text, state.get('filepath'), set_status, set_stats)

    export_btn = make_button(bottom, '  Экспортировать в Excel', btn_process,
                             primary=True, width=34)
    export_btn.pack(fill='x')

    win.mainloop()


if __name__ == '__main__':
    run_gui()
