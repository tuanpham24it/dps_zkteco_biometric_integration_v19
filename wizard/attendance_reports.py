# -*- coding: utf-8 -*-
from odoo import api, fields, models, _
from datetime import datetime, timedelta
import base64
import io
import pytz
import math
from dateutil.relativedelta import relativedelta
from odoo.tools import DEFAULT_SERVER_DATETIME_FORMAT
from pytz import timezone, UTC

# Customized by Tunn
import xlsxwriter
# from odoo.tools.misc import xlsxwriter

from collections import defaultdict
from odoo.exceptions import UserError, ValidationError


class EmployeeAttendanceReports(models.TransientModel):
    _name = 'employee.attendance.reports'
    _description = 'Employee Attendance Reports'

    report_type = fields.Selection([
        ('attendance_report', 'Attendance Report'),
        ('absence_report', 'Absence Report'),
        ('daily_summary_report', 'Daily Summary Report'),
        ('calculate_attendance_difference', 'Calculate Attendance Difference'),
    ], string="Report Type", default='attendance_report', required=True)

    employee_ids = fields.Many2many('hr.employee', string='Employee')
    start_date = fields.Date(string="Start Date", required=True)
    end_date = fields.Date(string="End Date", required=True)
    report_file_store = fields.Binary('File', readonly=True)
    report_file_name = fields.Char('Filename', readonly=True)

    def float_to_time_str(self, hours_float):
        total_minutes = int(round(hours_float * 60))
        h, m = divmod(total_minutes, 60)
        return f"{h:02d}:{m:02d}"

    def generate_report(self):
        for rec in self:
            employee_ids = rec.employee_ids or self.env['hr.employee'].search([])

            # ===============================================================
            # Attendance Report
            # ===============================================================
            if rec.report_type == 'attendance_report':
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})

                # Define styles
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#DCE6F1',
                    'border': 1, 'align': 'center'
                })
                cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                date_format = workbook.add_format({
                    'border': 1, 'num_format': 'yyyy-mm-dd', 'align': 'center'
                })
                total_format = workbook.add_format({
                    'bold': True, 'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'
                })

                # Iterate over employees
                for emp in employee_ids:
                    sheet_name = emp.name[:31] if emp.name else 'Employee'
                    worksheet = workbook.add_worksheet(sheet_name)

                    # Write headers
                    headers = [
                        'Sr No', 'Date', 'Employee Name', 'Barcode',
                        'Check In', 'Check Out', 'Worked Hours',
                        'Shortfall', 'Break Time', 'Overtime Hours'
                    ]
                    for col, header in enumerate(headers):
                        worksheet.write(0, col, header, header_format)

                    # Fetch attendance records
                    attendances = self.env['hr.attendance'].search([
                        ('employee_id', '=', emp.id),
                        ('check_in', '>=', rec.start_date),
                        ('check_in', '<=', rec.end_date)
                    ], order='check_in asc')

                    # Fill data
                    row = 1
                    total_worked = total_shortfall = total_break = total_overtime = 0.0

                    for idx, att in enumerate(attendances, start=1):
                        check_in = att.check_in or ''
                        check_out = att.check_out or ''
                        worked_hours = att.worked_hours or 0.0
                        shortfall = getattr(att, 'shortfall', 0.0) or 0.0
                        break_time = getattr(att, 'break_time', 0.0) or 0.0
                        overtime_hours = getattr(att, 'validated_overtime_hours', 0.0) or 0.0
                        if overtime_hours <= 0:
                            overtime_hours = 0.0

                        total_worked += worked_hours
                        total_shortfall += shortfall
                        total_break += break_time
                        total_overtime += overtime_hours

                        # Write data row
                        worksheet.write(row, 0, idx, cell_format)
                        worksheet.write(row, 1, check_in.date() if check_in else '', date_format)
                        worksheet.write(row, 2, emp.name or '', cell_format)
                        worksheet.write(row, 3, emp.barcode or '', cell_format)
                        worksheet.write(row, 4, str(check_in) if check_in else '', cell_format)
                        worksheet.write(row, 5, str(check_out) if check_out else '', cell_format)
                        worksheet.write(row, 6, self.float_to_time_str(worked_hours), cell_format)
                        worksheet.write(row, 7, self.float_to_time_str(shortfall), cell_format)
                        worksheet.write(row, 8, self.float_to_time_str(break_time), cell_format)
                        worksheet.write(row, 9, self.float_to_time_str(overtime_hours), cell_format)
                        row += 1

                    # Write totals row
                    if row > 1:
                        worksheet.write(row, 5, 'Total', total_format)
                        worksheet.write(row, 6, self.float_to_time_str(total_worked), total_format)
                        worksheet.write(row, 7, self.float_to_time_str(total_shortfall), total_format)
                        worksheet.write(row, 8, self.float_to_time_str(total_break), total_format)
                        worksheet.write(row, 9, self.float_to_time_str(total_overtime), total_format)

                    # Adjust column widths
                    worksheet.set_column('A:J', 18)

                workbook.close()

                # Prepare binary output
                output.seek(0)
                file_data = base64.b64encode(output.read())
                filename = f"Employee_Attendance_Report_{fields.Date.today()}.xlsx"

                rec.write({
                    'report_file_store': file_data,
                    'report_file_name': filename
                })

                # Return file download action
                return {
                    'type': 'ir.actions.act_url',
                    'url': f'/web/content/?model={rec._name}&id={rec.id}&field=report_file_store&filename_field=report_file_name&download=true',
                    'target': 'self',
                }

            # ===============================================================
            # Absence Report
            # ===============================================================
            elif rec.report_type == 'absence_report':
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})

                # Define styles
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#F4B084',
                    'border': 1, 'align': 'center'
                })
                cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                date_format = workbook.add_format({
                    'border': 1, 'num_format': 'yyyy-mm-dd', 'align': 'center'
                })
                total_format = workbook.add_format({
                    'bold': True, 'bg_color': '#FFF2CC',
                    'border': 1, 'align': 'center'
                })

                # Prepare date range
                start_date = fields.Date.from_string(rec.start_date)
                end_date = fields.Date.from_string(rec.end_date)
                date_range = [start_date + timedelta(days=i)
                              for i in range((end_date - start_date).days + 1)]

                # Employees
                employee_ids = rec.employee_ids or self.env['hr.employee'].search([])

                for emp in employee_ids:
                    # Create a new worksheet for each employee
                    sheet_name = emp.name[:31] if emp.name else 'Employee'
                    worksheet = workbook.add_worksheet(sheet_name)

                    # Write headers
                    headers = ['Sr No', 'Employee Name', 'Barcode', 'Date', 'Day', 'Reason (Absent)']
                    for col, header in enumerate(headers):
                        worksheet.write(0, col, header, header_format)

                    row = 1
                    sr_no = 1
                    total_absent = 0

                    for date in date_range:
                        # Check if attendance exists for that date
                        attendance = self.env['hr.attendance'].search([
                            ('employee_id', '=', emp.id),
                            ('check_in', '>=', datetime.combine(date, datetime.min.time())),
                            ('check_in', '<=', datetime.combine(date, datetime.max.time())),
                        ], limit=1)

                        # If no attendance found => mark as absent
                        if not attendance:
                            worksheet.write(row, 0, sr_no, cell_format)
                            worksheet.write(row, 1, emp.name or '', cell_format)
                            worksheet.write(row, 2, emp.barcode or '', cell_format)
                            worksheet.write(row, 3, date, date_format)
                            worksheet.write(row, 4, date.strftime('%A'), cell_format)  # Day name
                            worksheet.write(row, 5, 'Absent', cell_format)

                            sr_no += 1
                            total_absent += 1
                            row += 1

                    # Totals row for that employee
                    if total_absent > 0:
                        worksheet.write(row, 4, 'Total Absent Days', total_format)
                        worksheet.write(row, 5, total_absent, total_format)

                    worksheet.set_column('A:F', 20)

                # Finalize workbook
                workbook.close()

                # Prepare binary output
                output.seek(0)
                file_data = base64.b64encode(output.read())
                filename = f"Employee_Absence_Report_{fields.Date.today()}.xlsx"

                rec.write({
                    'report_file_store': file_data,
                    'report_file_name': filename
                })

                # Return file download action
                return {
                    'type': 'ir.actions.act_url',
                    'url': f'/web/content/?model={rec._name}&id={rec.id}&field=report_file_store&filename_field=report_file_name&download=true',
                    'target': 'self',
                }

            # ===============================================================
            # Daily Summary
            # ===============================================================
            elif rec.report_type == 'daily_summary_report':
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})

                # Define formats
                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#A9D08E',
                    'border': 1, 'align': 'center'
                })
                cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                date_time_format = workbook.add_format({
                    'border': 1, 'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center'
                })
                date_format = workbook.add_format({
                    'border': 1, 'num_format': 'yyyy-mm-dd', 'align': 'center'
                })

                # Prepare employees and date range
                employee_ids = rec.employee_ids or self.env['hr.employee'].search([])
                start_date = fields.Date.from_string(rec.start_date)
                end_date = fields.Date.from_string(rec.end_date)
                start_dt = datetime.combine(start_date, datetime.min.time())
                end_dt = datetime.combine(end_date, datetime.max.time())

                # Iterate per employee
                for emp in employee_ids:
                    sheet_name = emp.name[:31] if emp.name else 'Employee'
                    worksheet = workbook.add_worksheet(sheet_name)

                    # Write header row
                    headers = [
                        'Sr No', 'Employee Name', 'Barcode',
                        'Date', 'Day', 'Punch Time',
                        'Device', 'ZKTeco User ID',
                        'Status', 'Status Number'
                    ]
                    for col, header in enumerate(headers):
                        worksheet.write(0, col, header, header_format)

                    # Fetch logs from zkteco.device.logs
                    logs = self.env['zkteco.device.logs'].search([
                        ('employee_id', '=', emp.id),
                        ('user_punch_time', '>=', start_dt),
                        ('user_punch_time', '<=', end_dt),
                    ], order='user_punch_time asc')

                    row = 1
                    sr_no = 1

                    for log in logs:
                        punch_time = log.user_punch_time
                        date_val = punch_time.date() if punch_time else ''
                        day_name = punch_time.strftime('%A') if punch_time else ''

                        # Get status and numeric value safely
                        status = dict(log._fields['status'].selection).get(log.status, '') if log.status else ''
                        status_number = getattr(log, 'status_number', '')

                        worksheet.write(row, 0, sr_no, cell_format)
                        worksheet.write(row, 1, emp.name or '', cell_format)
                        worksheet.write(row, 2, emp.barcode or '', cell_format)
                        worksheet.write(row, 3, date_val, date_format)
                        worksheet.write(row, 4, day_name, cell_format)
                        worksheet.write(row, 5, punch_time, date_time_format)
                        worksheet.write(row, 6, getattr(log.device_id, 'name', ''), cell_format)
                        worksheet.write(row, 7, getattr(log, 'zketco_duser_id', ''), cell_format)
                        worksheet.write(row, 8, status, cell_format)
                        worksheet.write(row, 9, status_number, cell_format)

                        sr_no += 1
                        row += 1

                    if not logs:
                        worksheet.write(row, 0, 'No logs found for this employee.', cell_format)

                    worksheet.set_column('A:J', 20)

                workbook.close()

                # Prepare binary output
                output.seek(0)
                file_data = base64.b64encode(output.read())
                filename = f"Employee_Daily_Summary_Report_{fields.Date.today()}.xlsx"

                rec.write({
                    'report_file_store': file_data,
                    'report_file_name': filename
                })

                # Return file download action
                return {
                    'type': 'ir.actions.act_url',
                    'url': f'/web/content/?model={rec._name}&id={rec.id}&field=report_file_store&filename_field=report_file_name&download=true',
                    'target': 'self',
                }

            # ===============================================================
            # Attendance Difference (Placeholders)
            # ===============================================================
            elif rec.report_type == 'calculate_attendance_difference':
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})

                header_format = workbook.add_format({
                    'bold': True, 'bg_color': '#BDD7EE', 'border': 1, 'align': 'center'
                })

                cell_format = workbook.add_format({'border': 1, 'align': 'center'})
                date_format = workbook.add_format({
                    'border': 1, 'num_format': 'yyyy-mm-dd', 'align': 'center'
                })

                datetime_format = workbook.add_format({
                    'border': 1, 'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center'
                })

                bold_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})

                def float_to_time_str(hours_float):
                    """Convert float hours to HH:MM format"""
                    total_seconds = int(hours_float * 3600)
                    hrs = total_seconds // 3600
                    mins = (total_seconds % 3600) // 60
                    return f"{hrs:02d}:{mins:02d}"

                employee_ids = rec.employee_ids or self.env['hr.employee'].search([])
                start_date = fields.Date.from_string(rec.start_date)
                end_date = fields.Date.from_string(rec.end_date)
                start_dt = datetime.combine(start_date, datetime.min.time())
                end_dt = datetime.combine(end_date, datetime.max.time())

                for emp in employee_ids:
                    worksheet = workbook.add_worksheet(emp.name[:31] if emp.name else 'Employee')
                    headers = [
                        'Sr No', 'Employee Name', 'Barcode', 'Date', 'Day',
                        'Check In', 'Check Out', 'Actual Working Hours',
                        'Expected Working Hours', 'Difference (Hrs)', 'Status'
                    ]

                    for col, header in enumerate(headers):
                        worksheet.write(0, col, header, header_format)

                    expected_hours = emp.resource_calendar_id.hours_per_day or 0.0

                    attendances = self.env['hr.attendance'].search([
                        ('employee_id', '=', emp.id),
                        ('check_in', '>=', start_dt),
                        ('check_in', '<=', end_dt)
                    ], order='check_in asc')

                    row = 1
                    sr_no = 1
                    total_actual = 0.0
                    total_expected = 0.0
                    total_diff = 0.0
                    for att in attendances:
                        check_in = att.check_in
                        check_out = att.check_out
                        actual_hours = att.worked_hours or 0.0
                        difference = actual_hours - expected_hours
                        total_actual += actual_hours
                        total_expected += expected_hours
                        total_diff += difference
                        date_val = check_in.date() if check_in else ''
                        day_name = check_in.strftime('%A') if check_in else ''

                        if difference > 0.1:
                            status = 'Overtime'
                        elif difference < -0.1:
                            status = 'Less Hours'
                        else:
                            status = 'On Time'

                        worksheet.write(row, 0, sr_no, cell_format)
                        worksheet.write(row, 1, emp.name or '', cell_format)
                        worksheet.write(row, 2, emp.barcode or '', cell_format)
                        worksheet.write(row, 3, date_val, date_format)
                        worksheet.write(row, 4, day_name, cell_format)
                        worksheet.write(row, 5, check_in, datetime_format)
                        worksheet.write(row, 6, check_out, datetime_format)
                        worksheet.write(row, 7, float_to_time_str(actual_hours), cell_format)
                        worksheet.write(row, 8, float_to_time_str(expected_hours), cell_format)
                        worksheet.write(row, 9, float_to_time_str(difference), cell_format)
                        worksheet.write(row, 10, status, cell_format)

                        sr_no += 1
                        row += 1

                    if attendances:
                        worksheet.write(row, 6, 'Total', bold_format)
                        worksheet.write(row, 7, float_to_time_str(total_actual), bold_format)
                        worksheet.write(row, 8, float_to_time_str(total_expected), bold_format)
                        worksheet.write(row, 9, float_to_time_str(total_diff), bold_format)
                        worksheet.write(row, 10, '', bold_format)

                    else:
                        worksheet.write(1, 0, 'No attendance found for this employee.', cell_format)
                    worksheet.set_column('A:K', 20)

                workbook.close()
                output.seek(0)
                file_data = base64.b64encode(output.read())
                filename = f"Attendance_Difference_Report_{fields.Date.today()}.xlsx"

                rec.write({
                    'report_file_store': file_data,
                    'report_file_name': filename
                })

                return {
                    'type': 'ir.actions.act_url',
                    'url': f'/web/content/?model={rec._name}&id={rec.id}&field=report_file_store&filename_field=report_file_name&download=true',
                    'target': 'self',
                }


    def action_cancel(self):
            return {'type': 'ir.actions.act_window_close'}
