#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Расчёт фактического эффекта от индексации тарифов.
GUI на PyQt5, совместимо с запуском из IDLE.
"""

import sys
import os
import json
import time
import traceback
from datetime import datetime
from collections import defaultdict

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLabel, QComboBox, QPushButton, QCheckBox,
    QFileDialog, QTextEdit, QGroupBox, QRadioButton, QButtonGroup,
    QScrollArea, QFrame, QMessageBox, QProgressBar, QSplitter,
    QSizePolicy
)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QColor


# ─────────────────────────────────────────────
#  1. ЗАГРУЗКА МЭППИНГОВ
# ─────────────────────────────────────────────

def load_mappings():
    """Загрузка файла мэппингов из того же каталога, где лежит скрипт."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    mappings_path = os.path.join(script_dir, 'mappings.json')
    if not os.path.exists(mappings_path):
        # Попробуем текущую директорию
        mappings_path = 'mappings.json'
    with open(mappings_path, 'r', encoding='utf-8') as f:
        return json.load(f)


MAPPINGS = load_mappings()

ALL_MRF_FULL = MAPPINGS['ALL_MRF_FULL']
MRF_FULL_TO_SHORT = MAPPINGS['MRF_FULL_TO_SHORT']
MRF_SHORT_TO_FULL = MAPPINGS['MRF_SHORT_TO_FULL']
MRF_TO_RF = MAPPINGS['MRF_TO_RF']
RF_TO_MRF = MAPPINGS['RF_TO_MRF']
SERVICES = MAPPINGS['SERVICES']
SHORT_SEGMENTS = MAPPINGS['SHORT_SEGMENTS']
SUBSEGMENTS = MAPPINGS['SUBSEGMENTS']

MONTHS = [
    'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
    'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
]

# Заполнения для выделения
FILL_BLUE = PatternFill(start_color='99CCFF', end_color='99CCFF', fill_type='solid')
FILL_GREEN = PatternFill(start_color='92D050', end_color='92D050', fill_type='solid')
FILL_YELLOW = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')


# ─────────────────────────────────────────────
#  2. УТИЛИТЫ
# ─────────────────────────────────────────────

def find_column_ci(columns, target):
    """
   _case-insensitive_ поиск колонки по имени.
    Возвращает точное имя колонки или None.
    """
    target_lower = str(target).strip().lower()
    for c in columns:
        if str(c).strip().lower() == target_lower:
            return c
    return None


def find_column_contains_ci(columns, substrings):
    """
    Поиск колонки, содержащей ВСЕ указанные подстроки (case-insensitive).
    substrings — список строк.
    """
    subs = [s.strip().lower() for s in substrings]
    for c in columns:
        cl = str(c).strip().lower()
        if all(s in cl for s in subs):
            return c
    return None


def normalize_mrf_to_report(mrf_value):
    """
    Привести значение МРФ к формату отчёта.
    В отчёте: СЗ, ДВ — коротко; остальные — полностью.
    """
    if mrf_value is None:
        return None
    v = str(mrf_value).strip()
    # Если уже в формате отчёта
    if v in MRF_FULL_TO_SHORT.values():
        return v
    # Полное имя → короткое для СЗ и ДВ
    if v in MRF_FULL_TO_SHORT:
        return MRF_FULL_TO_SHORT[v]
    # Попробуем найти по частичному совпадению
    v_lower = v.lower()
    for full, short in MRF_FULL_TO_SHORT.items():
        if full.lower() == v_lower or short.lower() == v_lower:
            return short
    return v  # Вернём как есть


def normalize_mrf_to_full(mrf_value):
    """Привести к полному имени МРФ."""
    if mrf_value is None:
        return None
    v = str(mrf_value).strip()
    if v in MRF_SHORT_TO_FULL:
        return MRF_SHORT_TO_FULL[v]
    if v in MRF_FULL_TO_SHORT:
        return v
    v_lower = v.lower()
    for full_name in ALL_MRF_FULL:
        if full_name.lower() == v_lower:
            return full_name
    return v


def normalize_rf(rf_value):
    """Нормализовать название филиала (убрать лишние пробелы)."""
    if rf_value is None:
        return None
    return str(rf_value).strip()


def rf_to_report_mrf(rf_value):
    """По названию филиала определить МРФ в формате отчёта."""
    rf = normalize_rf(rf_value)
    if rf in RF_TO_MRF:
        full_mrf = RF_TO_MRF[rf]
        return MRF_FULL_TO_SHORT.get(full_mrf, full_mrf)
    return None


def rf_matches_mrf(rf_value, mrf_full_name):
    """Проверить, относится ли филиал к данному МРФ."""
    rf = normalize_rf(rf_value)
    # Прямое соответствие
    if rf in RF_TO_MRF and RF_TO_MRF[rf] == mrf_full_name:
        return True
    # Проверка через MRF_TO_RF
    rf_list = MRF_TO_RF.get(mrf_full_name, [])
    if rf in rf_list:
        return True
    # Попробуем нормализовать: убрать "филиал" и сравнить
    rf_lower = rf.lower().replace('филиал', '').strip()
    for r in rf_list:
        r_lower = r.lower().replace('филиал', '').strip()
        if rf_lower == r_lower:
            return True
    return False


def safe_float(val, default=0.0):
    """Безопасное преобразование в float."""
    if val is None:
        return default
    try:
        v = float(str(val).replace(',', '.').replace(' ', '').strip())
        return v
    except (ValueError, TypeError):
        return default


def timestamp_str():
    return datetime.now().strftime('%H:%M:%S')


# ─────────────────────────────────────────────
#  3. ДВИЖОК РАСЧЁТА
# ─────────────────────────────────────────────

class CalculationEngine:
    """Расчёт фактического эффекта для каждого абонента."""

    @staticmethod
    def calc_fact_effect(accrual_report, ap_before, ap_after):
        """
        Расчёт фактического эффекта для одного абонента.

        Parameters
        ----------
        accrual_report : float
            Начисления за отчётный месяц.
        ap_before : float
            АП до индексации (тариф до).
        ap_after : float
            АП после индексации (тариф после).

        Returns
        -------
        float
            Фактический эффект (в рублях без НДС).
        """
        if ap_before > 0:
            # Основной случай: max(0, начисления_отчёт - начисления_ДО)
            delta = accrual_report - ap_before
            return max(0.0, delta)
        else:
            # АП до == 0: специальный случай
            if abs(accrual_report - ap_after) < 0.01:
                # начисления == АП после → прогнозный эффект
                return max(0.0, ap_after - ap_before)
            elif accrual_report > ap_before:
                # начисления > АП до
                return max(0.0, accrual_report - ap_before)
            else:
                return 0.0

    @staticmethod
    def process_base(df_base, col_mapping, accruals_dict, mode='A',
                     filter_dates=None, filter_services=None,
                     filter_segments=None, filter_subsegments=None):
        """
        Обработать базу индексации и вернуть агрегированные результаты.

        Parameters
        ----------
        df_base : pd.DataFrame
            Объединённая база индексации.
        col_mapping : dict
            Маппинг колонок: {role: column_name}
            role может быть: 'msisdn', 'ap_before', 'ap_after', 'date',
                            'service', 'percent', 'mrf', 'rf',
                            'short_segment', 'subsegment',
                            'accrual_report' (для режима Б)
        accruals_dict : dict
            {MSISDN: начисления_за_отчётный_месяц} для режима А.
            Пустой dict для режима Б.
        mode : str
            'A' или 'B'.
        filter_dates : set or None
        filter_services : set or None
        filter_segments : set or None
        filter_subsegments : set or None

        Returns
        -------
        dict
            {(service, date, percent, mrf_report, rf, segment, subsegment): total_effect_rub}
        """
        results = defaultdict(float)
        stats = {'total_rows': 0, 'processed': 0, 'no_accrual': 0, 'zero_effect': 0}

        col_msisdn = col_mapping['msisdn']
        col_ap_before = col_mapping['ap_before']
        col_ap_after = col_mapping.get('ap_after', '')
        col_date = col_mapping.get('date', '')
        col_service = col_mapping.get('service', '')
        col_percent = col_mapping.get('percent', '')
        col_mrf = col_mapping.get('mrf', '')
        col_rf = col_mapping.get('rf', '')
        col_segment = col_mapping.get('short_segment', '')
        col_subsegment = col_mapping.get('subsegment', '')
        col_accrual = col_mapping.get('accrual_report', '')  # для режима Б

        for idx, row in df_base.iterrows():
            stats['total_rows'] += 1

            # Получаем значения
            msisdn = str(row.get(col_msisdn, '')).strip()
            if not msisdn:
                continue

            ap_before = safe_float(row.get(col_ap_before, 0))
            ap_after = safe_float(row.get(col_ap_after, 0)) if col_ap_after else 0.0

            date_val = str(row.get(col_date, '')).strip() if col_date else ''
            service_val = str(row.get(col_service, '')).strip() if col_service else ''
            percent_val = str(row.get(col_percent, '')).strip() if col_percent else ''
            mrf_val = str(row.get(col_mrf, '')).strip() if col_mrf else ''
            rf_val = str(row.get(col_rf, '')).strip() if col_rf else ''
            segment_val = str(row.get(col_segment, '')).strip() if col_segment else ''
            subsegment_val = str(row.get(col_subsegment, '')).strip() if col_subsegment else ''

            # Фильтры
            if filter_dates and date_val not in filter_dates:
                continue
            if filter_services and service_val not in filter_services:
                continue
            if filter_segments and segment_val not in filter_segments:
                continue
            if filter_subsegments and subsegment_val not in filter_subsegments:
                continue

            # Получаем начисления за отчётный месяц
            if mode == 'A':
                # Из кеша начислений
                msisdn_key = msisdn
                # Попробуем разные форматы MSISDN
                accrual_report = accruals_dict.get(msisdn_key)
                if accrual_report is None:
                    # Попробуем без 8 или 7
                    if msisdn.startswith('8') and len(msisdn) == 11:
                        accrual_report = accruals_dict.get('7' + msisdn[1:])
                    elif msisdn.startswith('7') and len(msisdn) == 11:
                        accrual_report = accruals_dict.get('8' + msisdn[1:])
                    if accrual_report is None:
                        # Попробуем только 10 цифр
                        if len(msisdn) >= 10:
                            accrual_report = accruals_dict.get(msisdn[-10:])
                if accrual_report is None:
                    accrual_report = 0.0
                    stats['no_accrual'] += 1
            else:
                # Режим Б — начисления в базе
                accrual_report = safe_float(row.get(col_accrual, 0)) if col_accrual else 0.0

            # Приводим МРФ к формату отчёта
            mrf_report = normalize_mrf_to_report(mrf_val)
            # Если не удалось, попробуем по РФ
            if not mrf_report and rf_val:
                mrf_report = rf_to_report_mrf(rf_val)
            if not mrf_report:
                mrf_report = mrf_val  # Оставим как есть

            # Расчёт эффекта
            effect = CalculationEngine.calc_fact_effect(accrual_report, ap_before, ap_after)

            if effect == 0:
                stats['zero_effect'] += 1

            # Агрегация
            key = (service_val, date_val, percent_val, mrf_report, rf_val, segment_val, subsegment_val)
            results[key] += effect
            stats['processed'] += 1

        return dict(results), stats


# ─────────────────────────────────────────────
#  4. ОБРАБОТКА ОТЧЁТА (openpyxl)
# ─────────────────────────────────────────────

class ReportHandler:
    """Работа с файлом отчёта через openpyxl."""

    def __init__(self, filepath):
        self.filepath = filepath
        self.wb = load_workbook(filepath)
        self.sheet_name = 'Факт-эффект индексации 2026'
        if self.sheet_name not in self.wb.sheetnames:
            # Попробуем найти похожий лист
            for name in self.wb.sheetnames:
                if 'факт' in name.lower() and 'индекс' in name.lower():
                    self.sheet_name = name
                    break
        self.ws = self.wb[self.sheet_name]
        self.header_row = 1
        self.col_map = {}  # {role: col_index_1based}
        self.target_col = None  # Колонка для записи ФАКТ за выбранный месяц
        self._parse_headers()

    def _parse_headers(self):
        """Разобрать заголовки и найти колонки."""
        headers = {}
        for col_idx in range(1, self.ws.max_column + 1):
            val = self.ws.cell(row=self.header_row, column=col_idx).value
            if val is not None:
                headers[col_idx] = str(val).strip()

        # Ищем ключевые колонки с приоритетом точного совпадения
        key_names = {
            'service': ['услуга'],
            'date': ['дата индексации', 'дата'],
            'percent': ['процент индексации', 'процент'],
            'mrf': ['мрф', 'макрорегион'],
            'rf': ['рф'],
            'short_segment': ['краткий сегмент', 'сегмент'],
            'subsegment': ['подсегмент'],
        }

        for role, search_terms in key_names.items():
            best_match = None
            best_score = -1

            for col_idx, header_text in headers.items():
                hl = header_text.lower()
                for term in search_terms:
                    tl = term.lower()
                    if tl in hl:
                        # Оценка: чем ближе длина заголовка к длине термина, тем лучше
                        # (предпочитаем точные совпадения)
                        score = 100 - abs(len(hl) - len(tl))
                        # Штраф за «мрф» при поиске «рф»
                        if role == 'rf' and 'мрф' in hl and 'м' not in tl:
                            score = -1  # очень низкий приоритет
                        if score > best_score:
                            best_score = score
                            best_match = col_idx
                        break  # первый совпавший термин

            if best_match is not None and best_score > 0:
                self.col_map[role] = best_match

    def find_target_column(self, month_name):
        """Найти колонку «ФАКТ Эффект ... [месяц] 2026»."""
        for col_idx in range(1, self.ws.max_column + 1):
            val = self.ws.cell(row=self.header_row, column=col_idx).value
            if val is None:
                continue
            header = str(val).strip().lower()
            if ('факт' in header and
                    month_name.lower() in header and
                    '2026' in header):
                self.target_col = col_idx
                return col_idx
        # Попробуем более мягкий поиск
        for col_idx in range(1, self.ws.max_column + 1):
            val = self.ws.cell(row=self.header_row, column=col_idx).value
            if val is None:
                continue
            header = str(val).strip().lower()
            if ('эффект' in header and
                    month_name.lower() in header and
                    '2026' in header):
                self.target_col = col_idx
                return col_idx
        return None

    def find_matching_row(self, key):
        """
        Найти строку в отчёте, соответствующую ключу.
        key = (service, date, percent, mrf_report, rf, segment, subsegment)
        Возвращает номер строки (1-based) или None.
        """
        service, date_val, percent_val, mrf_report, rf_val, segment, subsegment = key

        # Сопоставим с колонками отчёта
        col_service = self.col_map.get('service')
        col_date = self.col_map.get('date')
        col_percent = self.col_map.get('percent')
        col_mrf = self.col_map.get('mrf')
        col_rf = self.col_map.get('rf')
        col_segment = self.col_map.get('short_segment')
        col_subsegment = self.col_map.get('subsegment')

        for row_idx in range(2, self.ws.max_row + 1):
            match = True

            # Проверяем каждое поле
            checks = [
                (col_service, service),
                (col_date, date_val),
                (col_percent, percent_val),
                (col_mrf, mrf_report),
                (col_rf, rf_val),
                (col_segment, segment),
                (col_subsegment, subsegment),
            ]

            for col_idx, target_val in checks:
                if col_idx is None:
                    continue
                cell_val = self.ws.cell(row=row_idx, column=col_idx).value
                if cell_val is None:
                    cell_str = ''
                else:
                    cell_str = str(cell_val).strip()

                target_str = str(target_val).strip() if target_val else ''

                # Гибкое сравнение
                if not self._values_match(cell_str, target_str):
                    match = False
                    break

            if match:
                return row_idx

        return None

    def _values_match(self, cell_val, target_val):
        """Гибкое сравнение значений ячеек."""
        if not cell_val and not target_val:
            return True
        if not cell_val or not target_val:
            return False

        cv = cell_val.strip().lower()
        tv = target_val.strip().lower()

        # Прямое совпадение
        if cv == tv:
            return True

        # Числовое сравнение (для процентов)
        try:
            cv_num = float(cv.replace(',', '.').replace('%', '').replace(' ', ''))
            tv_num = float(tv.replace(',', '.').replace('%', '').replace(' ', ''))
            if abs(cv_num - tv_num) < 0.001:
                return True
        except (ValueError, TypeError):
            pass

        # Для МРФ: полное ↔ короткое
        if cv in MRF_FULL_TO_SHORT and MRF_FULL_TO_SHORT[cv] == tv:
            return True
        if tv in MRF_FULL_TO_SHORT and MRF_FULL_TO_SHORT[tv] == cv:
            return True
        if cv in MRF_SHORT_TO_FULL and MRF_SHORT_TO_FULL[cv].lower() == tv:
            return True
        if tv in MRF_SHORT_TO_FULL and MRF_SHORT_TO_FULL[tv].lower() == cv:
            return True

        # Частичное совпадение для названий услуг/регионов
        if len(cv) > 3 and len(tv) > 3:
            if cv in tv or tv in cv:
                return True

        return False

    def write_results(self, results_dict, month_name, log_callback=None):
        """
        Записать результаты в отчёт.

        Parameters
        ----------
        results_dict : dict
            {key: effect_rub} — эффект в рублях.
        month_name : str
            Название месяца.
        log_callback : callable or None

        Returns
        -------
        dict
            Статистика записи.
        """
        target_col = self.find_target_column(month_name)
        if target_col is None:
            raise ValueError(
                f"Не найдена колонка «ФАКТ Эффект от индексации тыс.руб. без НДС {month_name} 2026» в отчёте."
            )

        stats = {'written': 0, 'added_rows': 0, 'updated_rows': 0, 'yellow_cells': 0}

        for key, effect_rub in results_dict.items():
            effect_thousands = effect_rub / 1000.0  # В тыс.руб.

            row_idx = self.find_matching_row(key)

            if row_idx is not None:
                # Строка найдена — перезаписываем
                cell = self.ws.cell(row=row_idx, column=target_col)
                cell.value = round(effect_thousands, 2)
                stats['updated_rows'] += 1

                if effect_thousands > 0:
                    cell.fill = FILL_GREEN
                    stats['written'] += 1
                else:
                    cell.fill = FILL_YELLOW
                    stats['yellow_cells'] += 1
            else:
                # Добавить новую строку
                new_row = self.ws.max_row + 1
                self._write_new_row(new_row, key, target_col, effect_thousands)
                stats['added_rows'] += 1
                if effect_thousands > 0:
                    stats['written'] += 1

            if log_callback:
                service, date_val, percent_val, mrf_report, rf_val, seg, subseg = key
                log_callback(
                    f"  Записано: {service} | {rf_val} | {seg} | {subseg} | "
                    f"эффект={effect_thousands:.2f} тыс.руб."
                )

        return stats

    def _write_new_row(self, row_idx, key, target_col, effect_thousands):
        """Записать новую строку в отчёт."""
        service, date_val, percent_val, mrf_report, rf_val, segment, subsegment = key

        field_map = {
            'service': service,
            'date': date_val,
            'percent': percent_val,
            'mrf': mrf_report,
            'rf': rf_val,
            'short_segment': segment,
            'subsegment': subsegment,
        }

        for role, value in field_map.items():
            col_idx = self.col_map.get(role)
            if col_idx is not None:
                cell = self.ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                cell.fill = FILL_BLUE

        # Записать эффект
        cell = self.ws.cell(row=row_idx, column=target_col)
        cell.value = round(effect_thousands, 2)
        if effect_thousands > 0:
            cell.fill = FILL_GREEN
        else:
            cell.fill = FILL_BLUE

    def save(self):
        """Сохранить отчёт."""
        self.wb.save(self.filepath)

    def close(self):
        """Закрыть workbook."""
        self.wb.close()


# ─────────────────────────────────────────────
#  5. GUI — ГЛАВНОЕ ОКНО
# ─────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Расчёт фактического эффекта от индексации тарифов')
        self.setMinimumSize(1100, 850)

        # Состояние
        self.report_path = None
        self.report_handler = None
        self.selected_month = None
        self.mode = 'A'  # 'A' или 'B'
        self.selected_mrf_list = []

        # Per-MRF state
        self.current_mrf_idx = 0
        self.current_mrf_full = None

        # Base data
        self.df_base = None
        self.base_columns = []
        self.base_file_paths = []

        # Accruals cache
        self.accruals_cache = {}  # {msisdn: value}
        self.accruals_file_name = ''
        self.accruals_col_mapping = {}

        # Column mapping
        self.col_mapping = {}

        # Filters
        self.available_dates = []
        self.available_services = []
        self.available_segments = []
        self.available_subsegments = []

        # Log file
        self.log_file = None
        log_dir = os.path.dirname(os.path.abspath(__file__))
        log_path = os.path.join(log_dir, 'indexation_log.txt')
        try:
            self.log_file = open(log_path, 'a', encoding='utf-8')
        except:
            pass

        self._build_ui()
        self._update_visibility()
        self._log('Программа запущена.')

    def _build_ui(self):
        """Построить интерфейс."""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # Scroll area for top controls
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # === Группа 1: Глобальные настройки ===
        grp_global = QGroupBox('1. Глобальные настройки')
        lay_global = QGridLayout()

        lay_global.addWidget(QLabel('Отчётный месяц:'), 0, 0)
        self.cmb_month = QComboBox()
        self.cmb_month.addItems(MONTHS)
        self.cmb_month.setCurrentIndex(0)
        lay_global.addWidget(self.cmb_month, 0, 1)

        lay_global.addWidget(QLabel('Режим:'), 0, 2)
        mode_layout = QHBoxLayout()
        self.rb_mode_a = QRadioButton('А (отдельный файл начислений)')
        self.rb_mode_b = QRadioButton('Б (начисления в базе)')
        self.rb_mode_a.setChecked(True)
        self.rb_mode_a.toggled.connect(self._on_mode_changed)
        mode_layout.addWidget(self.rb_mode_a)
        mode_layout.addWidget(self.rb_mode_b)
        lay_global.addLayout(mode_layout, 0, 3)

        # Отчёт
        lay_global.addWidget(QLabel('Файл отчёта:'), 1, 0)
        self.lbl_report_file = QLabel('не загружен')
        self.lbl_report_file.setStyleSheet('color: gray;')
        lay_global.addWidget(self.lbl_report_file, 1, 1, 1, 2)
        btn_report = QPushButton('Выбрать отчёт...')
        btn_report.clicked.connect(self._load_report)
        lay_global.addWidget(btn_report, 1, 3)

        grp_global.setLayout(lay_global)
        scroll_layout.addWidget(grp_global)

        # === Группа 2: Выбор МРФ ===
        grp_mrf = QGroupBox('2. Макрорегионы (отметьте нужные)')
        lay_mrf = QGridLayout()
        self.chk_mrf = {}
        for i, mrf in enumerate(ALL_MRF_FULL):
            cb = QCheckBox(mrf)
            self.chk_mrf[mrf] = cb
            lay_mrf.addWidget(cb, i // 4, i % 4)

        btn_select_all_mrf = QPushButton('Выбрать все')
        btn_select_all_mrf.clicked.connect(self._select_all_mrf)
        lay_mrf.addWidget(btn_select_all_mrf, 2, 0)
        btn_deselect_all_mrf = QPushButton('Снять все')
        btn_deselect_all_mrf.clicked.connect(self._deselect_all_mrf)
        lay_mrf.addWidget(btn_deselect_all_mrf, 2, 1)

        grp_mrf.setLayout(lay_mrf)
        scroll_layout.addWidget(grp_mrf)

        # === Группа 3: Обработка МРФ ===
        grp_process = QGroupBox('3. Обработка МРФ — загрузка баз и начислений')
        lay_process = QGridLayout()

        self.lbl_current_mrf = QLabel('МРФ не выбран')
        self.lbl_current_mrf.setStyleSheet('font-weight: bold; font-size: 13px;')
        lay_process.addWidget(self.lbl_current_mrf, 0, 0, 1, 4)

        # База индексации
        self.btn_start_mrf = QPushButton('▶ Начать обработку выбранных МРФ')
        self.btn_start_mrf.clicked.connect(self._start_mrf_processing)
        self.btn_start_mrf.setStyleSheet('background-color: #4CAF50; color: white; padding: 8px; font-size: 13px;')
        lay_process.addWidget(self.btn_start_mrf, 1, 0, 1, 4)

        # Загрузка базы
        lay_process.addWidget(QLabel('База индексации:'), 2, 0)
        self.lbl_base_files = QLabel('не загружена')
        self.lbl_base_files.setWordWrap(True)
        lay_process.addWidget(self.lbl_base_files, 2, 1, 1, 2)
        self.btn_load_base = QPushButton('Загрузить базу...')
        self.btn_load_base.clicked.connect(self._load_base)
        self.btn_load_base.setEnabled(False)
        lay_process.addWidget(self.btn_load_base, 2, 3)

        self.btn_load_more_base = QPushButton('Загрузить ещё базу?')
        self.btn_load_more_base.clicked.connect(self._load_more_base)
        self.btn_load_more_base.setEnabled(False)
        lay_process.addWidget(self.btn_load_more_base, 3, 3)

        # Кеш начислений
        self.lbl_cache_status = QLabel('Кеш: пуст')
        self.lbl_cache_status.setStyleSheet('color: gray;')
        lay_process.addWidget(self.lbl_cache_status, 4, 0, 1, 2)
        self.btn_load_accruals = QPushButton('Загрузить начисления...')
        self.btn_load_accruals.clicked.connect(self._load_accruals)
        self.btn_load_accruals.setEnabled(False)
        lay_process.addWidget(self.btn_load_accruals, 4, 2)
        self.btn_clear_cache = QPushButton('Очистить кеш')
        self.btn_clear_cache.clicked.connect(self._clear_cache)
        lay_process.addWidget(self.btn_clear_cache, 4, 3)

        grp_process.setLayout(lay_process)
        scroll_layout.addWidget(grp_process)

        # === Группа 4: Маппинг колонок ===
        grp_mapping = QGroupBox('4. Маппинг колонок базы индексации')
        lay_mapping = QGridLayout()

        self.mapping_roles = [
            ('msisdn', 'MSISDN (уник. идентификатор)'),
            ('ap_before', 'Начисления ДО индексации (АП до)'),
            ('ap_after', 'Начисления ПОСЛЕ индексации (АП после)'),
            ('date', 'Дата индексации'),
            ('service', 'Услуга'),
            ('percent', 'Процент индексации'),
            ('mrf', 'МРФ'),
            ('rf', 'РФ (филиал)'),
            ('short_segment', 'Краткий сегмент'),
            ('subsegment', 'Подсегмент'),
        ]

        self.cmb_mapping = {}
        for i, (role, label) in enumerate(self.mapping_roles):
            lay_mapping.addWidget(QLabel(label + ':'), i, 0)
            cmb = QComboBox()
            cmb.setMinimumWidth(250)
            self.cmb_mapping[role] = cmb
            lay_mapping.addWidget(cmb, i, 1)

        # Для режима Б — дополнительная колонка начислений
        self.lbl_accrual_col = QLabel('Начисления за отчётный месяц (режим Б):')
        self.cmb_accrual_col_b = QComboBox()
        self.cmb_accrual_col_b.setMinimumWidth(250)
        lay_mapping.addWidget(self.lbl_accrual_col, len(self.mapping_roles), 0)
        lay_mapping.addWidget(self.cmb_accrual_col_b, len(self.mapping_roles), 1)
        self.lbl_accrual_col.setVisible(False)
        self.cmb_accrual_col_b.setVisible(False)

        # Маппинг колонок файла начислений (режим А)
        self.lbl_sep_accrual = QLabel('─── Маппинг колонок файла начислений ───')
        self.lbl_sep_accrual.setStyleSheet('font-weight: bold;')
        row_offset = len(self.mapping_roles) + 1
        lay_mapping.addWidget(self.lbl_sep_accrual, row_offset, 0, 1, 2)

        lay_mapping.addWidget(QLabel('MSISDN в файле начислений:'), row_offset + 1, 0)
        self.cmb_accrual_msisdn = QComboBox()
        self.cmb_accrual_msisdn.setMinimumWidth(250)
        lay_mapping.addWidget(self.cmb_accrual_msisdn, row_offset + 1, 1)

        lay_mapping.addWidget(QLabel('Начисления за отчётный месяц:'), row_offset + 2, 0)
        self.cmb_accrual_value = QComboBox()
        self.cmb_accrual_value.setMinimumWidth(250)
        lay_mapping.addWidget(self.cmb_accrual_value, row_offset + 2, 1)

        self.btn_rebuild_cache = QPushButton('Пересобрать кеш')
        self.btn_rebuild_cache.clicked.connect(self._rebuild_cache_from_ui)
        lay_mapping.addWidget(self.btn_rebuild_cache, row_offset + 3, 0, 1, 2)

        grp_mapping.setLayout(lay_mapping)
        scroll_layout.addWidget(grp_mapping)

        # === Группа 5: Фильтры ===
        grp_filters = QGroupBox('5. Фильтры (отметьте нужные значения)')
        lay_filters = QVBoxLayout()

        # Даты
        lay_dates = QHBoxLayout()
        lay_dates.addWidget(QLabel('Даты индексации:'))
        self.filter_dates_layout = QHBoxLayout()
        self.chk_dates = {}
        lay_dates.addLayout(self.filter_dates_layout)
        lay_filters.addLayout(lay_dates)

        # Услуги
        lay_services = QHBoxLayout()
        lay_services.addWidget(QLabel('Услуги:'))
        self.filter_services_layout = QHBoxLayout()
        self.chk_services = {}
        lay_services.addLayout(self.filter_services_layout)
        lay_filters.addLayout(lay_services)

        # Краткие сегменты
        lay_segments = QHBoxLayout()
        lay_segments.addWidget(QLabel('Краткие сегменты:'))
        self.filter_segments_layout = QHBoxLayout()
        self.chk_segments = {}
        lay_segments.addLayout(self.filter_segments_layout)
        lay_filters.addLayout(lay_segments)

        # Подсегменты
        lay_subsegments = QHBoxLayout()
        lay_subsegments.addWidget(QLabel('Подсегменты:'))
        self.filter_subsegments_layout = QHBoxLayout()
        self.chk_subsegments = {}
        lay_subsegments.addLayout(self.filter_subsegments_layout)
        lay_filters.addLayout(lay_subsegments)

        # Кнопки выбора всех фильтров
        lay_filter_btns = QHBoxLayout()
        btn_select_all_filters = QPushButton('Выбрать все фильтры')
        btn_select_all_filters.clicked.connect(self._select_all_filters)
        lay_filter_btns.addWidget(btn_select_all_filters)
        btn_deselect_all_filters = QPushButton('Снять все фильтры')
        btn_deselect_all_filters.clicked.connect(self._deselect_all_filters)
        lay_filter_btns.addWidget(btn_deselect_all_filters)
        lay_filters.addLayout(lay_filter_btns)

        grp_filters.setLayout(lay_filters)
        scroll_layout.addWidget(grp_filters)

        # === Кнопка расчёта ===
        self.btn_calculate = QPushButton('🔢 РАССЧИТАТЬ ФАКТИЧЕСКИЙ ЭФФЕКТ')
        self.btn_calculate.clicked.connect(self._calculate)
        self.btn_calculate.setEnabled(False)
        self.btn_calculate.setStyleSheet(
            'background-color: #2196F3; color: white; padding: 12px; '
            'font-size: 15px; font-weight: bold;'
        )
        scroll_layout.addWidget(self.btn_calculate)

        # === Кнопка следующий МРФ ===
        lay_nav = QHBoxLayout()
        self.btn_next_mrf = QPushButton('Следующий МРФ ▶')
        self.btn_next_mrf.clicked.connect(self._next_mrf)
        self.btn_next_mrf.setEnabled(False)
        self.btn_next_mrf.setStyleSheet('background-color: #FF9800; color: white; padding: 8px;')
        lay_nav.addWidget(self.btn_next_mrf)

        self.btn_save_report = QPushButton('💾 Сохранить отчёт')
        self.btn_save_report.clicked.connect(self._save_report)
        self.btn_save_report.setEnabled(False)
        self.btn_save_report.setStyleSheet('background-color: #9C27B0; color: white; padding: 8px;')
        lay_nav.addWidget(self.btn_save_report)
        scroll_layout.addLayout(lay_nav)

        # === Журнал ===
        grp_log = QGroupBox('Журнал операций')
        lay_log = QVBoxLayout()
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumHeight(200)
        self.txt_log.setStyleSheet('font-family: Consolas, monospace; font-size: 11px;')
        lay_log.addWidget(self.txt_log)
        grp_log.setLayout(lay_log)
        scroll_layout.addWidget(grp_log)

        scroll.setWidget(scroll_widget)
        main_layout.addWidget(scroll)

    # ─── Обработчики событий ───

    def _on_mode_changed(self):
        if self.rb_mode_a.isChecked():
            self.mode = 'A'
            self._log('Режим переключён: А (отдельный файл начислений)')
        else:
            self.mode = 'B'
            self._log('Режим переключён: Б (начисления в базе индексации)')
        self._clear_cache()
        self._update_visibility()

    def _update_visibility(self):
        """Обновить видимость элементов в зависимости от режима."""
        is_a = (self.mode == 'A')
        self.btn_load_accruals.setVisible(is_a)
        self.lbl_cache_status.setVisible(is_a)
        self.btn_clear_cache.setVisible(is_a)
        self.lbl_sep_accrual.setVisible(is_a)
        self.cmb_accrual_msisdn.parentWidget().setVisible(True)
        # Для режима Б показываем дополнительную колонку
        self.lbl_accrual_col.setVisible(not is_a)
        self.cmb_accrual_col_b.setVisible(not is_a)

    def _select_all_mrf(self):
        for cb in self.chk_mrf.values():
            cb.setChecked(True)

    def _deselect_all_mrf(self):
        for cb in self.chk_mrf.values():
            cb.setChecked(False)

    def _select_all_filters(self):
        for d in [self.chk_dates, self.chk_services, self.chk_segments, self.chk_subsegments]:
            for cb in d.values():
                cb.setChecked(True)

    def _deselect_all_filters(self):
        for d in [self.chk_dates, self.chk_services, self.chk_segments, self.chk_subsegments]:
            for cb in d.values():
                cb.setChecked(False)

    def _load_report(self):
        """Загрузить файл отчёта."""
        path, _ = QFileDialog.getOpenFileName(
            self, 'Выберите файл отчёта', '',
            'Excel файлы (*.xlsx *.xls);;Все файлы (*)'
        )
        if not path:
            return
        try:
            self.report_path = path
            if self.report_handler:
                self.report_handler.close()
            self.report_handler = ReportHandler(path)
            fname = os.path.basename(path)
            self.lbl_report_file.setText(fname)
            self.lbl_report_file.setStyleSheet('color: green; font-weight: bold;')
            self._log(f'✓ Отчёт загружен: {fname}')
            self._log(f'  Лист: {self.report_handler.sheet_name}')
            self._log(f'  Найдены колонки: {list(self.report_handler.col_map.keys())}')

            target = self.report_handler.find_target_column(
                self.cmb_month.currentText()
            )
            if target:
                self._log(f'  Целевая колонка (ФАКТ): #{target}')
            else:
                self._log(f'  ⚠ Целевая колонка для месяца '
                          f'«{self.cmb_month.currentText()}» не найдена.')
        except Exception as e:
            self._log(f'✗ Ошибка загрузки отчёта: {e}')
            QMessageBox.critical(self, 'Ошибка', f'Не удалось загрузить отчёт:\n{e}')

    def _get_selected_mrf_list(self):
        """Получить список выбранных МРФ."""
        return [mrf for mrf, cb in self.chk_mrf.items() if cb.isChecked()]

    def _start_mrf_processing(self):
        """Начать обработку первого выбранного МРФ."""
        if self.report_handler is None:
            QMessageBox.warning(self, 'Внимание', 'Сначала загрузите файл отчёта.')
            return

        self.selected_mrf_list = self._get_selected_mrf_list()
        if not self.selected_mrf_list:
            QMessageBox.warning(self, 'Внимание', 'Выберите хотя бы один макрорегион.')
            return

        self.current_mrf_idx = 0
        self._activate_current_mrf()

    def _activate_current_mrf(self):
        """Активировать обработку текущего МРФ."""
        if self.current_mrf_idx >= len(self.selected_mrf_list):
            self._log('✓ Все МРФ обработаны!')
            self.lbl_current_mrf.setText('Все МРФ обработаны!')
            self.btn_calculate.setEnabled(False)
            self.btn_load_base.setEnabled(False)
            self.btn_load_accruals.setEnabled(False)
            self.btn_next_mrf.setEnabled(False)
            self.btn_save_report.setEnabled(True)
            return

        self.current_mrf_full = self.selected_mrf_list[self.current_mrf_idx]
        short = MRF_FULL_TO_SHORT.get(self.current_mrf_full, self.current_mrf_full)
        self.lbl_current_mrf.setText(
            f'МРФ: {self.current_mrf_full} ({short}) '
            f'[{self.current_mrf_idx + 1}/{len(self.selected_mrf_list)}]'
        )

        # Сброс состояния для нового МРФ
        self.df_base = None
        self.base_columns = []
        self.base_file_paths = []
        self.accruals_cache = {}
        self.accruals_file_name = ''
        self.col_mapping = {}

        # Очистить GUI
        self.lbl_base_files.setText('не загружена')
        self.lbl_cache_status.setText('Кеш: пуст')
        self.lbl_cache_status.setStyleSheet('color: gray;')
        for cmb in self.cmb_mapping.values():
            cmb.clear()
        self.cmb_accrual_col_b.clear()
        self.cmb_accrual_msisdn.clear()
        self.cmb_accrual_value.clear()
        self._clear_filter_checkboxes()

        # Включить кнопки
        self.btn_load_base.setEnabled(True)
        self.btn_load_more_base.setEnabled(False)
        if self.mode == 'A':
            self.btn_load_accruals.setEnabled(True)
        self.btn_calculate.setEnabled(False)
        self.btn_next_mrf.setEnabled(False)

        self._log(f'▶ Начинаем обработку МРФ: {self.current_mrf_full}')

    def _clear_filter_checkboxes(self):
        """Удалить все чекбоксы фильтров."""
        for layout in [self.filter_dates_layout, self.filter_services_layout,
                       self.filter_segments_layout, self.filter_subsegments_layout]:
            while layout.count():
                item = layout.takeAt(0)
                w = item.widget()
                if w:
                    w.deleteLater()
        self.chk_dates = {}
        self.chk_services = {}
        self.chk_segments = {}
        self.chk_subsegments = {}

    def _load_base(self):
        """Загрузить файл базы индексации."""
        path, _ = QFileDialog.getOpenFileName(
            self, f'База индексации для МРФ «{self.current_mrf_full}»', '',
            'Excel файлы (*.xlsx *.xls);;Все файлы (*)'
        )
        if not path:
            return

        try:
            self._log(f'Загрузка базы: {os.path.basename(path)}...')
            QApplication.processEvents()

            df = pd.read_excel(path, sheet_name='Лист1', engine='openpyxl')
            self._log(f'  Загружено строк: {len(df)}, колонок: {len(df.columns)}')

            if self.df_base is not None:
                self.df_base = pd.concat([self.df_base, df], ignore_index=True)
            else:
                self.df_base = df

            self.base_file_paths.append(path)
            self.base_columns = list(self.df_base.columns)

            # Обновить GUI
            names = [os.path.basename(p) for p in self.base_file_paths]
            self.lbl_base_files.setText(', '.join(names))
            self.lbl_base_files.setStyleSheet('color: green;')

            # Заполнить выпадающие списки колонок
            self._populate_column_combos()

            # Извлечь фильтры
            self._populate_filters()

            # Включить кнопку «ещё базу»
            self.btn_load_more_base.setEnabled(True)

            self._log(f'✓ База загружена: {os.path.basename(path)} '
                      f'(всего строк в объединённой базе: {len(self.df_base)})')

            # Проверить готовность к расчёту
            self._check_ready_to_calculate()

        except Exception as e:
            self._log(f'✗ Ошибка загрузки базы: {e}')
            traceback.print_exc()
            QMessageBox.critical(self, 'Ошибка', f'Не удалось загрузить базу:\n{e}')

    def _load_more_base(self):
        """Загрузить дополнительную базу для того же МРФ."""
        self._load_base()

    def _populate_column_combos(self):
        """Заполнить выпадающие списки колонок из базы."""
        cols = self.base_columns
        for role, cmb in self.cmb_mapping.items():
            cmb.clear()
            cmb.addItem('— не выбрано —')
            cmb.addItems([str(c) for c in cols])

            # Автовыбор по названию
            auto = self._auto_select_column(role, cols)
            if auto:
                idx = cmb.findText(str(auto))
                if idx >= 0:
                    cmb.setCurrentIndex(idx)

        # Для режима Б
        self.cmb_accrual_col_b.clear()
        self.cmb_accrual_col_b.addItem('— не выбрано —')
        self.cmb_accrual_col_b.addItems([str(c) for c in cols])

    def _auto_select_column(self, role, columns):
        """Автоматически подобрать колонку по роли."""
        search_map = {
            'msisdn': ['msisdn', 'мсисден', 'номер'],
            'ap_before': ['ап до', 'до индексации', 'тариф до', 'начисления до'],
            'ap_after': ['ап после', 'после индексации', 'тариф после', 'начисления после'],
            'date': ['дата индексации', 'дата'],
            'service': ['услуга', 'сервис', 'вид услуги'],
            'percent': ['процент', '% индексации'],
            'mrf': ['мрф', 'макрорегион'],
            'rf': ['рф', 'филиал', 'регион'],
            'short_segment': ['краткий сегмент', 'сегмент'],
            'subsegment': ['подсегмент'],
        }

        terms = search_map.get(role, [])
        for term in terms:
            result = find_column_contains_ci(columns, [term])
            if result:
                return result
        return None

    def _populate_filters(self):
        """Заполнить чекбоксы фильтров из данных базы."""
        self._clear_filter_checkboxes()
        if self.df_base is None:
            return

        # Определяем колонки (используя автовыбор или выбранные)
        col_date = self._get_selected_col('date')
        col_service = self._get_selected_col('service')
        col_segment = self._get_selected_col('short_segment')
        col_subsegment = self._get_selected_col('subsegment')

        # Даты
        if col_date:
            dates = sorted(self.df_base[col_date].dropna().unique())
            for d in dates:
                d_str = str(d).strip()
                if d_str and d_str != 'nan':
                    cb = QCheckBox(d_str)
                    cb.setChecked(True)
                    self.chk_dates[d_str] = cb
                    self.filter_dates_layout.addWidget(cb)

        # Услуги
        if col_service:
            services = sorted(self.df_base[col_service].dropna().unique())
            for s in services:
                s_str = str(s).strip()
                if s_str and s_str != 'nan':
                    cb = QCheckBox(s_str)
                    cb.setChecked(True)
                    self.chk_services[s_str] = cb
                    self.filter_services_layout.addWidget(cb)
        else:
            # Используем эталонный список
            for s in SERVICES:
                cb = QCheckBox(s)
                cb.setChecked(True)
                self.chk_services[s] = cb
                self.filter_services_layout.addWidget(cb)

        # Краткие сегменты
        if col_segment:
            segs = sorted(self.df_base[col_segment].dropna().unique())
            for s in segs:
                s_str = str(s).strip()
                if s_str and s_str != 'nan':
                    cb = QCheckBox(s_str)
                    cb.setChecked(True)
                    self.chk_segments[s_str] = cb
                    self.filter_segments_layout.addWidget(cb)
        else:
            for s in SHORT_SEGMENTS:
                cb = QCheckBox(s)
                cb.setChecked(True)
                self.chk_segments[s] = cb
                self.filter_segments_layout.addWidget(cb)

        # Подсегменты
        if col_subsegment:
            subs = sorted(self.df_base[col_subsegment].dropna().unique())
            for s in subs:
                s_str = str(s).strip()
                if s_str and s_str != 'nan':
                    cb = QCheckBox(s_str)
                    cb.setChecked(True)
                    self.chk_subsegments[s_str] = cb
                    self.filter_subsegments_layout.addWidget(cb)
        else:
            for s in SUBSEGMENTS:
                cb = QCheckBox(s)
                cb.setChecked(True)
                self.chk_subsegments[s] = cb
                self.filter_subsegments_layout.addWidget(cb)

    def _get_selected_col(self, role):
        """Получить выбранное имя колонки для роли."""
        cmb = self.cmb_mapping.get(role)
        if cmb and cmb.currentIndex() > 0:
            return cmb.currentText()
        return None

    def _load_accruals(self):
        """Загрузить файл начислений (режим А)."""
        if self.mode != 'A':
            return

        path, _ = QFileDialog.getOpenFileName(
            self, 'Файл начислений (режим А)', '',
            'Excel файлы (*.xlsx *.xls);;Все файлы (*)'
        )
        if not path:
            return

        try:
            self._log(f'Загрузка начислений: {os.path.basename(path)}...')
            self.lbl_cache_status.setText('Кеш: загрузка...')
            self.lbl_cache_status.setStyleSheet('color: orange;')
            QApplication.processEvents()
            t0 = time.time()

            df_acc = pd.read_excel(path, sheet_name='Sheet', engine='openpyxl')
            self._df_accruals = df_acc  # Сохраняем для пересборки кеша
            self._log(f'  Начисления загружены: {len(df_acc)} строк за {time.time()-t0:.1f} сек')

            # Заполнить колонки начислений
            self.cmb_accrual_msisdn.clear()
            self.cmb_accrual_value.clear()
            acc_cols = list(df_acc.columns)
            self.cmb_accrual_msisdn.addItems([str(c) for c in acc_cols])
            self.cmb_accrual_value.addItems([str(c) for c in acc_cols])

            # Автовыбор MSISDN
            msisdn_col = find_column_contains_ci(acc_cols, ['msisdn', 'мсисден', 'номер'])
            if msisdn_col:
                idx = self.cmb_accrual_msisdn.findText(str(msisdn_col))
                if idx >= 0:
                    self.cmb_accrual_msisdn.setCurrentIndex(idx)

            # Автовыбор колонки начислений (ищем что-то с "начисл" или "сумма" или "фаクト")
            month = self.cmb_month.currentText()
            accrual_col = find_column_contains_ci(acc_cols, ['факт', month.lower()])
            if accrual_col is None:
                accrual_col = find_column_contains_ci(acc_cols, ['начисл', month.lower()])
            if accrual_col is None:
                accrual_col = find_column_contains_ci(acc_cols, ['начисл'])
            if accrual_col is None:
                # Берём последнюю числовую колонку
                for c in reversed(acc_cols):
                    if df_acc[c].dtype in ['float64', 'int64', 'float32']:
                        accrual_col = c
                        break
            if accrual_col:
                idx = self.cmb_accrual_value.findText(str(accrual_col))
                if idx >= 0:
                    self.cmb_accrual_value.setCurrentIndex(idx)

            # Построить кеш {MSISDN: начисления}
            msisdn_col_name = self.cmb_accrual_msisdn.currentText()
            value_col_name = self.cmb_accrual_value.currentText()

            if (not msisdn_col_name or msisdn_col_name == '— не выбрано —' or
                    not value_col_name or value_col_name == '— не выбрано —'):
                self._log('⚠ Выберите колонки MSISDN и начислений для файла начислений!')
                # Всё равно построим кеш, пользователь сможет пересобрать позже
                return

            self._build_accruals_cache(df_acc, msisdn_col_name, value_col_name)

            fname = os.path.basename(path)
            self.accruals_file_name = fname
            self.lbl_cache_status.setText(f'✓ Кеш: {fname} ({len(self.accruals_cache)} записей)')
            self.lbl_cache_status.setStyleSheet('color: green; font-weight: bold;')
            self._log(f'✓ Кеш начислений сформирован: {len(self.accruals_cache)} абонентов')

            self._check_ready_to_calculate()

        except Exception as e:
            self._log(f'✗ Ошибка загрузки начислений: {e}')
            traceback.print_exc()
            QMessageBox.critical(self, 'Ошибка', f'Не удалось загрузить начисления:\n{e}')

    def _build_accruals_cache(self, df_acc, msisdn_col, value_col):
        """Построить словарь кеша начислений."""
        self.accruals_cache = {}
        for _, row in df_acc.iterrows():
            msisdn = str(row.get(msisdn_col, '')).strip()
            if msisdn and msisdn != 'nan':
                val = safe_float(row.get(value_col, 0))
                self.accruals_cache[msisdn] = val

    def _rebuild_cache_from_ui(self):
        """Пересобрать кеш при изменении выбора колонок начислений."""
        if not hasattr(self, '_df_accruals') or self._df_accruals is None:
            self._log('⚠ Нет загруженных начислений для пересборки кеша.')
            return

        msisdn_col = self.cmb_accrual_msisdn.currentText()
        value_col = self.cmb_accrual_value.currentText()

        if (not msisdn_col or msisdn_col == '— не выбрано —' or
                not value_col or value_col == '— не выбрано —'):
            self._log('⚠ Выберите колонки MSISDN и начислений.')
            return

        try:
            self._build_accruals_cache(self._df_accruals, msisdn_col, value_col)
            self.lbl_cache_status.setText(
                f'✓ Кех: {self.accruals_file_name} ({len(self.accruals_cache)} записей)'
            )
            self.lbl_cache_status.setStyleSheet('color: green; font-weight: bold;')
            self._log(f'✓ Кех пересобран: {len(self.accruals_cache)} абонентов')
            self._check_ready_to_calculate()
        except Exception as e:
            self._log(f'✗ Ошибка пересборки кеша: {e}')

    def _rebuild_accruals_cache(self):
        """Пересобрать кеш при изменении выбора колонок."""
        if self.mode != 'A':
            return
        # Нужно перечитать файл — но мы не храним DataFrame.
        # Вместо этого покажем предупреждение.
        self._log('⚠ Для пересборки кеша загрузите файл начислений заново.')

    def _clear_cache(self):
        """Очистить кеш начислений."""
        self.accruals_cache = {}
        self.accruals_file_name = ''
        self.lbl_cache_status.setText('Кеш: пуст')
        self.lbl_cache_status.setStyleSheet('color: gray;')
        self._log('Кеш начислений очищен.')
        self._check_ready_to_calculate()

    def _check_ready_to_calculate(self):
        """Проверить, можно ли запустить расчёт."""
        ready = True
        reasons = []

        if self.df_base is None or len(self.df_base) == 0:
            ready = False
            reasons.append('не загружена база индексации')

        if self.mode == 'A' and len(self.accruals_cache) == 0:
            ready = False
            reasons.append('не загружены начисления (режим А)')

        # Проверим, что MSISDN выбран
        if not self._get_selected_col('msisdn'):
            ready = False
            reasons.append('не выбрана колонка MSISDN')

        if not self._get_selected_col('ap_before'):
            ready = False
            reasons.append('не выбрана колонка АП до')

        self.btn_calculate.setEnabled(ready)
        if ready:
            self._log('✓ Готово к расчёту!')
        elif reasons:
            self._log(f'  Не готово к расчёту: {"; ".join(reasons)}')

    def _get_filter_values(self, chk_dict):
        """Получить выбранные значения из чекбоксов. None если все выбраны или пусто."""
        checked = [k for k, cb in chk_dict.items() if cb.isChecked()]
        if len(checked) == 0:
            return None  # Нет выбранных — не фильтровать
        if len(checked) == len(chk_dict):
            return None  # Все выбраны — не фильтровать
        return set(checked)

    def _calculate(self):
        """Запустить расчёт фактического эффекта."""
        if self.report_handler is None:
            QMessageBox.warning(self, 'Внимание', 'Загрузите файл отчёта.')
            return

        try:
            self._log('')
            self._log('=' * 50)
            self._log(f'▶ Расчёт фактического эффекта для МРФ: {self.current_mrf_full}')
            t0 = time.time()

            # Собрать маппинг колонок
            col_mapping = {}
            for role, cmb in self.cmb_mapping.items():
                if cmb.currentIndex() > 0:
                    col_mapping[role] = cmb.currentText()
                else:
                    col_mapping[role] = ''

            # Режим Б — дополнительная колонка
            if self.mode == 'B':
                if self.cmb_accrual_col_b.currentIndex() > 0:
                    col_mapping['accrual_report'] = self.cmb_accrual_col_b.currentText()
                else:
                    self._log('✗ Не выбрана колонка начислений за отчётный месяц (режим Б)')
                    return

            # Проверить обязательные колонки
            if not col_mapping.get('msisdn'):
                self._log('✗ Не выбрана колонка MSISDN в базе индексации')
                return
            if not col_mapping.get('ap_before'):
                self._log('✗ Не выбрана колонка АП до в базе индексации')
                return

            # Фильтры
            filter_dates = self._get_filter_values(self.chk_dates)
            filter_services = self._get_filter_values(self.chk_services)
            filter_segments = self._get_filter_values(self.chk_segments)
            filter_subsegments = self._get_filter_values(self.chk_subsegments)

            # Фильтр по МРФ — только текущий
            # (фильтруем строки базы по выбранному МРФ)
            df_filtered = self._filter_base_by_mrf(self.df_base, col_mapping)
            self._log(f'  Строк в базе после фильтрации по МРФ: {len(df_filtered)} из {len(self.df_base)}')

            QApplication.processEvents()

            # Расчёт
            results, stats = CalculationEngine.process_base(
                df_base=df_filtered,
                col_mapping=col_mapping,
                accruals_dict=self.accruals_cache if self.mode == 'A' else {},
                mode=self.mode,
                filter_dates=filter_dates,
                filter_services=filter_services,
                filter_segments=filter_segments,
                filter_subsegments=filter_subsegments,
            )

            self._log(f'  Обработано строк: {stats["processed"]} из {stats["total_rows"]}')
            self._log(f'  Не найдено начислений: {stats["no_accrual"]}')
            self._log(f'  Нулевой эффект: {stats["zero_effect"]}')
            self._log(f'  Уникальных комбинаций: {len(results)}')

            if len(results) == 0:
                self._log('⚠ Нет данных для записи. Проверьте фильтры и маппинг колонок.')
                return

            # Запись в отчёт
            month_name = self.cmb_month.currentText()
            write_stats = self.report_handler.write_results(
                results, month_name, log_callback=self._log
            )

            elapsed = time.time() - t0
            self._log(f'')
            self._log(f'✓ Расчёт завершён за {elapsed:.1f} сек.')
            self._log(f'  Обновлено строк: {write_stats["updated_rows"]}')
            self._log(f'  Добавлено строк: {write_stats["added_rows"]}')
            self._log(f'  Зелёных (эффект > 0): {write_stats["written"]}')
            self._log(f'  Жёлтых (нет данных): {write_stats["yellow_cells"]}')

            # Общий итог
            total_effect = sum(results.values()) / 1000.0
            self._log(f'  ИТОГО эффект: {total_effect:.2f} тыс.руб. без НДС')

            self.btn_next_mrf.setEnabled(True)
            self.btn_save_report.setEnabled(True)

        except Exception as e:
            self._log(f'✗ Ошибка расчёта: {e}')
            traceback.print_exc()
            QMessageBox.critical(self, 'Ошибка расчёта', f'{e}\n\n{traceback.format_exc()}')

    def _filter_base_by_mrf(self, df, col_mapping):
        """Отфильтровать базу по текущему МРФ."""
        col_mrf = col_mapping.get('mrf', '')
        col_rf = col_mapping.get('rf', '')

        mrf_full = self.current_mrf_full
        mrf_short = MRF_FULL_TO_SHORT.get(mrf_full, mrf_full)
        rf_list = MRF_TO_RF.get(mrf_full, [])

        # Создаём маску
        mask = pd.Series([False] * len(df), index=df.index)

        if col_mrf and col_mrf in df.columns:
            mrf_vals = df[col_mrf].astype(str).str.strip()
            # Совпадение по полному имени или короткому коду
            mask = mask | (mrf_vals.str.lower() == mrf_full.lower())
            mask = mask | (mrf_vals.str.lower() == mrf_short.lower())

        if col_rf and col_rf in df.columns:
            rf_vals = df[col_rf].astype(str).str.strip()
            # Совпадение по списку филиалов
            rf_lower_set = set(r.lower() for r in rf_list)
            rf_mask = rf_vals.str.lower().isin(rf_lower_set)
            mask = mask | rf_mask

        # Если маска полностью False, попробуем более мягкий поиск
        if not mask.any():
            self._log(f'  ⚠ Прямое совпадение МРФ не найдено. '
                      f'Используем все строки базы.')
            return df  # Вернём все строки

        return df[mask].copy()

    def _next_mrf(self):
        """Перейти к следующему МРФ."""
        # Очистить кеш
        self._clear_cache()
        # Следующий МРФ
        self.current_mrf_idx += 1
        self._activate_current_mrf()

    def _save_report(self):
        """Сохранить файл отчёта."""
        if self.report_handler is None:
            return
        try:
            self.report_handler.save()
            self._log(f'✓ Отчёт сохранён: {self.report_path}')
            QMessageBox.information(self, 'Сохранено',
                                    f'Отчёт сохранён:\n{self.report_path}')
        except Exception as e:
            self._log(f'✗ Ошибка сохранения: {e}')
            QMessageBox.critical(self, 'Ошибка', f'Не удалось сохранить:\n{e}')

    def _log(self, message):
        """Записать сообщение в журнал."""
        ts = timestamp_str()
        line = f'[{ts}] {message}'
        self.txt_log.append(line)
        # Записать в файл
        if self.log_file:
            try:
                self.log_file.write(line + '\n')
                self.log_file.flush()
            except:
                pass
        # Прокрутить вниз
        sb = self.txt_log.verticalScrollBar()
        sb.setValue(sb.maximum())
        QApplication.processEvents()

    def closeEvent(self, event):
        """Обработка закрытия окна."""
        if self.report_handler:
            self.report_handler.close()
        if self.log_file:
            self.log_file.close()
        event.accept()


# ─────────────────────────────────────────────
#  6. ТОЧКА ВХОДА
# ─────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
