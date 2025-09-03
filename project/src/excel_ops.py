import openpyxl
from datetime import datetime, date
from calendar import monthrange
import uuid
import os
import time  # for filesystem flush on some platforms

# --------------------------------------------------------------------------- #
# Date helpers for multi-month support                                       #
# --------------------------------------------------------------------------- #

def _iter_month_ranges(start_date: date, end_date: date):
    """
    Yield (m_start, m_end) pairs where each pair covers the portion of the
    original [start_date, end_date] that falls inside a single month.

    The first span starts at `start_date`, the last span ends at `end_date`.
    """
    cur_year, cur_month = start_date.year, start_date.month

    while True:
        last_day = monthrange(cur_year, cur_month)[1]
        span_start = start_date if (cur_year == start_date.year and cur_month == start_date.month) \
            else date(cur_year, cur_month, 1)

        # decide span_end
        if cur_year == end_date.year and cur_month == end_date.month:
            span_end = end_date
        else:
            span_end = date(cur_year, cur_month, last_day)

        yield span_start, span_end

        # break if we just yielded the final month
        if cur_year == end_date.year and cur_month == end_date.month:
            break

        # advance month
        if cur_month == 12:
            cur_year += 1
            cur_month = 1
        else:
            cur_month += 1


def _get_month_sheet_name(target_date):
    """Get the sheet name for a given date in 'yy年M月' format."""
    year_short = str(target_date.year)[-2:]  # Get last 2 digits of year
    month = target_date.month
    return f"{year_short}年{month}月"


def _find_device_row(sheet, device_name):
    """Find the row index for a device name in column B."""
    for row in range(9, sheet.max_row + 1):
        if sheet.cell(row=row, column=2).value == device_name:
            return row
    return None


def _get_date_column(sheet, day):
    """Find the column index for a specific day in row 8."""
    for col in range(3, sheet.max_column + 1):
        if sheet.cell(row=8, column=col).value == day:
            return col
    return None


def _get_date_columns(sheet, start_date, end_date):
    """Get list of column indices for a date range."""
    columns = []
    for day in range(start_date.day, end_date.day + 1):
        col = _get_date_column(sheet, day)
        if col is None:
            raise ValueError(f"日付が見つかりません: {day}")
        columns.append(col)
    return columns


def _generate_booking_id():
    """Generate a unique booking ID."""
    return str(uuid.uuid4())[:8]


# --------------------------------------------------------------------------- #
# Utility iterator to enumerate device rows                                  #
# --------------------------------------------------------------------------- #

def _iter_device_rows(sheet):
    """
    Yield (row_index, device_name) tuples for each device registered
    in column B starting from row 9.
    """
    for row in range(9, sheet.max_row + 1):
        name = sheet.cell(row=row, column=2).value
        if name:
            yield row, str(name)



def check_availability(excel_path, device_name, start_date, end_date):
    """
    Check if a device is available for the specified date range.
    
    Args:
        excel_path: Path to the Excel file
        device_name: Name of the device to check
        start_date: Start date (inclusive)
        end_date: End date (inclusive)
        
    Returns:
        bool: True if available, False if already booked
    """
    wb = openpyxl.load_workbook(excel_path)
    try:
        for m_start, m_end in _iter_month_ranges(start_date, end_date):
            sheet_name = _get_month_sheet_name(m_start)
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"シートが見つかりません: {sheet_name}")

            sheet = wb[sheet_name]
            device_row = _find_device_row(sheet, device_name)
            if device_row is None:
                raise ValueError(f"デモ機が見つかりません: {device_name}")

            cols = _get_date_columns(sheet, m_start, m_end)
            for col in cols:
                if sheet.cell(row=device_row, column=col).value is not None:
                    return False
        return True
    finally:
        wb.close()


def find_available_device(excel_path, device_type, start_date, end_date):
    """
    Find the first available device of a given *type* for the specified period.

    Args:
        excel_path (str): Path to workbook
        device_type (str): Prefix of the device name, e.g., \"FE\" or \"PC\"
        start_date (date): Start date (inclusive)
        end_date (date): End date (inclusive)

    Returns:
        str | None: The first available device name or None if none found.
    """
    wb = openpyxl.load_workbook(excel_path)
    try:
        # collect candidate names from the start-month sheet
        start_sheet_name = _get_month_sheet_name(start_date)
        if start_sheet_name not in wb.sheetnames:
            raise ValueError(f"シートが見つかりません: {start_sheet_name}")

        start_sheet = wb[start_sheet_name]
        candidates = [
            name for _row, name in _iter_device_rows(start_sheet)
            if str(name).startswith(device_type)
        ]

        for dev_name in candidates:
            ok = True
            for m_start, m_end in _iter_month_ranges(start_date, end_date):
                sheet_name = _get_month_sheet_name(m_start)
                if sheet_name not in wb.sheetnames:
                    raise ValueError(f"シートが見つかりません: {sheet_name}")
                sheet = wb[sheet_name]
                row = _find_device_row(sheet, dev_name)
                if row is None:
                    raise ValueError(f"デモ機が見つかりません: {dev_name}")
                cols = _get_date_columns(sheet, m_start, m_end)
                if any(sheet.cell(row=row, column=c).value is not None for c in cols):
                    ok = False
                    break
            if ok:
                return dev_name
        return None
    finally:
        wb.close()


def book(excel_path, device_name, start_date, end_date, user_info):
    """
    Book a device for the specified date range.
    
    Args:
        excel_path: Path to the Excel file
        device_name: Name of the device to book
        start_date: Start date (inclusive)
        end_date: End date (inclusive)
        user_info: Dictionary with user information (name, extension, employee_id)
        
    Returns:
        str: Booking ID
        
    Raises:
        ValueError: If the device is already booked for any day in the range
    """
    # First check if available
    if not check_availability(excel_path, device_name, start_date, end_date):
        raise ValueError(f"指定された期間にデモ機 '{device_name}' は既に予約されています")
    
    wb = openpyxl.load_workbook(excel_path)
    
    # Get the correct sheet for the month
    sheet_name = _get_month_sheet_name(start_date)
    sheet = wb[sheet_name]
    
    # Find the device row
    device_row = _find_device_row(sheet, device_name)
    
    # Get columns for the date range
    date_columns = _get_date_columns(sheet, start_date, end_date)
    
    # Generate booking ID
    booking_id = _generate_booking_id()

    # Mark cells across all months
    booking_marker = f"C:{booking_id}"
    for m_start, m_end in _iter_month_ranges(start_date, end_date):
        sheet_name = _get_month_sheet_name(m_start)
        if sheet_name not in wb.sheetnames:
            wb.close()
            raise ValueError(f"シートが見つかりません: {sheet_name}")
        sheet_m = wb[sheet_name]
        device_row_m = _find_device_row(sheet_m, device_name)
        if device_row_m is None:
            wb.close()
            raise ValueError(f"デモ機が見つかりません: {device_name}")
        for col in _get_date_columns(sheet_m, m_start, m_end):
            sheet_m.cell(row=device_row_m, column=col, value=booking_marker)
    
    # Add entry to booking log
    log_sheet = wb['予約ログ']
    next_row = log_sheet.max_row + 1
    
    # Get current datetime as string
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    
    # Format dates as strings
    start_date_str = start_date.strftime("%Y-%m-%d")
    end_date_str = end_date.strftime("%Y-%m-%d")
    
    # Add log entry
    log_sheet.cell(row=next_row, column=1, value=booking_id)  # 予約ID
    log_sheet.cell(row=next_row, column=2, value=now)  # 予約日時
    log_sheet.cell(row=next_row, column=3, value=user_info.get('name', ''))  # 予約者名
    log_sheet.cell(row=next_row, column=4, value=user_info.get('extension', ''))  # 内線番号
    log_sheet.cell(row=next_row, column=5, value=user_info.get('employee_id', ''))  # 職番
    log_sheet.cell(row=next_row, column=6, value=device_name)  # デモ機名
    log_sheet.cell(row=next_row, column=7, value=start_date_str)  # 予約開始日
    log_sheet.cell(row=next_row, column=8, value=end_date_str)  # 予約終了日
    log_sheet.cell(row=next_row, column=9, value='予約中')  # ステータス
    
    # Save the workbook
    wb.save(excel_path)
    wb.close()
    
    return booking_id


def cancel(excel_path, booking_id):
    """
    Cancel a booking by ID.
    
    Args:
        excel_path: Path to the Excel file
        booking_id: ID of the booking to cancel
        
    Raises:
        ValueError: If the booking ID is not found
    """
    wb = openpyxl.load_workbook(excel_path)
    try:
        # --- locate booking log row ---
        log_sheet = wb['予約ログ']
        booking_row = None
        for row in range(2, log_sheet.max_row + 1):
            if log_sheet.cell(row=row, column=1).value == booking_id:
                booking_row = row
                break
        if booking_row is None:
            raise ValueError(f"予約IDが見つかりません: {booking_id}")

        device_name = log_sheet.cell(row=booking_row, column=6).value
        start_date = datetime.strptime(
            log_sheet.cell(row=booking_row, column=7).value, "%Y-%m-%d"
        ).date()
        end_date = datetime.strptime(
            log_sheet.cell(row=booking_row, column=8).value, "%Y-%m-%d"
        ).date()

        # --- clear booking marks across all months ---
        for m_start, m_end in _iter_month_ranges(start_date, end_date):
            sheet_name = _get_month_sheet_name(m_start)
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"シートが見つかりません: {sheet_name}")
            sheet = wb[sheet_name]
            device_row = _find_device_row(sheet, device_name)
            if device_row is None:
                raise ValueError(f"デモ機が見つかりません: {device_name}")
            # ハード化: 見出し行(8行目)を走査し、対象日のセルでかつ
            # この予約IDでマークされたセルのみをクリア
            marker_prefix = f"C:{booking_id}"
            for col in range(3, sheet.max_column + 1):
                header_val = sheet.cell(row=8, column=col).value
                if (
                    isinstance(header_val, int)
                    and m_start.day <= header_val <= m_end.day
                ):
                    cell_val = sheet.cell(row=device_row, column=col).value
                    if isinstance(cell_val, str) and cell_val.startswith(marker_prefix):
                        sheet.cell(row=device_row, column=col, value=None)

        wb.save(excel_path)
    finally:
        wb.close()

    # --- verification reopen pass (safety) ---
    wb_v = openpyxl.load_workbook(excel_path)
    try:
        for m_start, m_end in _iter_month_ranges(start_date, end_date):
            sheet_name = _get_month_sheet_name(m_start)
            if sheet_name in wb_v.sheetnames:
                sheet_v = wb_v[sheet_name]
                row_v = _find_device_row(sheet_v, device_name)
                if row_v is not None:
                    marker_prefix = f"C:{booking_id}"
                    for col in range(3, sheet_v.max_column + 1):
                        header_val = sheet_v.cell(row=8, column=col).value
                        if (
                            isinstance(header_val, int)
                            and m_start.day <= header_val <= m_end.day
                        ):
                            cell_val = sheet_v.cell(row=row_v, column=col).value
                            if (
                                isinstance(cell_val, str)
                                and cell_val.startswith(marker_prefix)
                            ):
                                sheet_v.cell(row=row_v, column=col, value=None)

        # ------------------------------------------------------------------ #
        # Fallback sweep: row-wide purge of any residual cells with this ID   #
        # ------------------------------------------------------------------ #
        for m_start, m_end in _iter_month_ranges(start_date, end_date):
            sheet_name = _get_month_sheet_name(m_start)
            if sheet_name in wb_v.sheetnames:
                sheet_v = wb_v[sheet_name]
                row_v = _find_device_row(sheet_v, device_name)
                if row_v is not None:
                    marker_prefix = f"C:{booking_id}"
                    for col in range(3, sheet_v.max_column + 1):
                        cell_val = sheet_v.cell(row=row_v, column=col).value
                        if isinstance(cell_val, str) and cell_val.startswith(marker_prefix):
                            sheet_v.cell(row=row_v, column=col, value=None)
        wb_v.save(excel_path)
        # give filesystem a moment to flush on Windows environments
        time.sleep(0.05)
    finally:
        wb_v.close()
