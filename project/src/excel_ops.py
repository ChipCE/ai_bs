import openpyxl
from datetime import datetime, date
import uuid
import os


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
    if start_date.month != end_date.month or start_date.year != end_date.year:
        raise ValueError("予約期間は同じ月内である必要があります")
    
    wb = openpyxl.load_workbook(excel_path)
    
    # Get the correct sheet for the month
    sheet_name = _get_month_sheet_name(start_date)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"シートが見つかりません: {sheet_name}")
    
    sheet = wb[sheet_name]
    
    # Find the device row
    device_row = _find_device_row(sheet, device_name)
    if device_row is None:
        raise ValueError(f"デモ機が見つかりません: {device_name}")
    
    # Get columns for the date range
    date_columns = _get_date_columns(sheet, start_date, end_date)
    
    # Check if any cell in the range is already booked
    for col in date_columns:
        cell_value = sheet.cell(row=device_row, column=col).value
        if cell_value is not None:
            return False
    
    wb.close()
    return True


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

    # Mark cells as booked with unique marker "C:<booking_id>"
    booking_marker = f"C:{booking_id}"
    for col in date_columns:
        sheet.cell(row=device_row, column=col, value=booking_marker)
    
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
    
    # Find booking in log
    log_sheet = wb['予約ログ']
    booking_row = None
    
    for row in range(2, log_sheet.max_row + 1):
        if log_sheet.cell(row=row, column=1).value == booking_id:
            booking_row = row
            break
    
    if booking_row is None:
        wb.close()
        raise ValueError(f"予約IDが見つかりません: {booking_id}")
    
    # Get booking details
    device_name = log_sheet.cell(row=booking_row, column=6).value
    start_date_str = log_sheet.cell(row=booking_row, column=7).value
    end_date_str = log_sheet.cell(row=booking_row, column=8).value
    
    # Parse dates
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d").date()
    
    # Get the correct sheet for the month
    sheet_name = _get_month_sheet_name(start_date)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"シートが見つかりません: {sheet_name}")
    
    sheet = wb[sheet_name]
    
    # Find the device row
    device_row = _find_device_row(sheet, device_name)
    if device_row is None:
        wb.close()
        raise ValueError(f"デモ機が見つかりません: {device_name}")
    
    # Get columns for the date range
    # 指示に従い _get_date_columns を使わず直接計算
    date_columns = [day + 2 for day in range(start_date.day, end_date.day + 1)]
    
    # Clear cells
    for col in date_columns:
        # 無条件でセルをクリア
        sheet.cell(row=device_row, column=col, value=None)
    
    # Update status in log (optional, could also add a new "キャンセル" entry)
    # log_sheet.cell(row=booking_row, column=9, value='キャンセル')
    
    # Save once and close
    wb.save(excel_path)
    wb.close()

    # --- double-check persistence: reopen, verify & clear again ---
    wb_verify = openpyxl.load_workbook(excel_path)
    if sheet_name in wb_verify.sheetnames:
        sheet_v = wb_verify[sheet_name]
        # Re-run clear just in case previous save was interrupted
        for col in date_columns:
            # 検証用ワークブックでも無条件でセルをクリア
            sheet_v.cell(row=device_row, column=col, value=None)
        wb_verify.save(excel_path)
    wb_verify.close()
