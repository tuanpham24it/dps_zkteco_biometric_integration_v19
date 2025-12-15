# -*- coding: utf-8 -*-
from odoo import api, fields, models, _
from datetime import datetime, timedelta
import base64
import os
import pytz
import math
from dateutil.relativedelta import relativedelta
from odoo.tools import DEFAULT_SERVER_DATETIME_FORMAT
from pytz import timezone, UTC
import io

# Customized by Tunn
import xlsxwriter
# from odoo.tools.misc import xlsxwriter

from collections import defaultdict
from odoo.exceptions import UserError, ValidationError


class EmployeeAttendanceReport(models.TransientModel):
    """
    Transient wizard model to generate Employee Attendance Reports.

    Users can select the source of the report (Attendance or Log),
    specify a date range, choose employees, and generate an Excel file.
    """
    _name = 'employee.attendance.report'
    _description = 'Employee Attendance Report Wizard'

    attendance_report_format = fields.Selection(
        [('attend', 'From Attendance'),
         ('log', 'From Log')],
        string='Report',
        required=True,
        default='attend'
    )

    report_date_start_from = fields.Datetime(
        'From',
        required=True,
        default=lambda self: datetime(2025, 2, 1, 0, 0, 0)
    )
    report_date_end_to = fields.Datetime(
        'To',
        required=True,
        default=datetime.today()
    )

    employee_attendance_report_file_store = fields.Binary(
        'File',
        readonly=True
    )
    employee_attendance_report_name = fields.Text(
        string='File Name'
    )
    is_printed = fields.Boolean(
        'Printed',
        default=False
    )

    employee_ids = fields.Many2many(
        'hr.employee',
        string='Employee'
    )

    attendance_excel_sheet_name = fields.Selection(
        [('badge_id', 'Badge Id'),
         ('emp_name', 'Name')],
        string='Sheet name',
        required=True,
        default='emp_name'
    )
    ################################################################################################################


    @api.onchange('attendance_report_format')
    def onchange_attendance_report_format(self):
        """
        Onchange method for 'attendance_report_format' field.

        Purpose:
        --------
        Automatically updates `report_date_start_from` and `report_date_end_to` fields based on the selected
        report format whenever the `attendance_report_format` field changes.

        Logic:
        ------
        - Retrieves the current day of the month.
        - Calculates `report_date_start_from` and `report_date_end_to` for the previous day:
            * `report_date_start_from` = Previous day at 00:00:00 hours.
            * `report_date_end_to`   = Previous day at 23:59:59 hours.
        - Updates `self.report_date_start_from` and `self.report_date_end_to` in string format (YYYY-MM-DD HH:MM:SS).

        Notes:
        ------
        - Uses `relativedelta` to adjust the current date safely.
        - Designed for dynamic report generation scenarios where users choose a report format.

        Exceptions:
        -----------
        Any unexpected error during computation will be caught, and a professional
        message will be raised to the user.
        """
        try:
            today = datetime.today()
            day = today.day

            report_date_start_from = today + relativedelta(day=day - 1, hour=0, minute=0, second=0)
            report_date_end_to = today + relativedelta(day=day - 1, hour=23, minute=59, second=59)

            self.report_date_start_from = report_date_start_from.strftime("%Y-%m-%d %H:%M:%S")
            self.report_date_end_to = report_date_end_to.strftime("%Y-%m-%d %H:%M:%S")

        except Exception as e:
            raise UserError(
                _("An error occurred while updating the attendance report date range. Please try again or contact support. Details: %s") % str(
                    e))

    def print_employee_attendance_in_excel(self, fl=None):
        """
        Export Attendance Data to XLSX File.

        Purpose:
        --------
        Generates an Excel report of attendance data based on the selected report format:
            - 'log'     : Exports raw attendance logs.
            - 'attend'  : Exports summarized attendance records.

        Workflow:
        ---------
        1. Determine report type from `attendance_report_format`.
        2. Convert `report_date_start_from` and `report_date_end_to` to the user's timezone using `new_timezone()`.
        3. Build a domain to filter records:
            - For 'log': Fetch `zkteco.device.logs` entries.
            - For 'attend': Fetch `hr.attendance` entries and group them by employee.
        4. Generate XLSX file content using helper methods:
            - `export_employee_attendance_from_logs()` for logs.
            - `add_employee_attendance_data()` for attendance records.
        5. Encode the file, update context, and return a dictionary to render a form view.

        Parameters:
        -----------
        fl : tuple or None
            A placeholder for the generated file content and filename. Defaults to an empty string.

        Returns:
        --------
        dict:
            An Odoo action dictionary to display the generated report file in a popup form.

        Exceptions:
        -----------
        Raises a professional `UserError` if any unexpected error occurs during the export process.
        """

        try:
            if fl is None:
                fl = ''

            if self.attendance_report_format == 'log':
                report_date_start_from = self.new_timezone(self.report_date_start_from)
                report_date_end_to = self.new_timezone(self.report_date_end_to)

                domain = [('user_punch_time', '>=', report_date_start_from),
                          ('user_punch_time', '<=', report_date_end_to)]

                if self.employee_ids:
                    domain.append(('employee_id', 'in', self.employee_ids.ids))

                attendance_logs = self.env['zkteco.device.logs'].search(domain)

                fl = self.export_employee_attendance_from_logs(attendance_logs)

            elif self.attendance_report_format == 'attend':
                report_date_start_from = self.new_timezone(self.report_date_start_from)
                report_date_end_to = self.new_timezone(self.report_date_end_to)

                domain = [
                    '|',
                    '&', ('check_in', '>=', report_date_start_from), ('check_out', '<=', report_date_end_to),
                    '&', '&', ('check_in', '>=', report_date_start_from), ('check_in', '<=', report_date_end_to), ('check_out', '=', False)
                ]

                if self.employee_ids:
                    domain.extend([('employee_id', 'in', self.employee_ids.ids)])

                attendances = self.env['hr.attendance'].search(domain)

                grouped_by_employee = defaultdict(list)
                for record in attendances:
                    grouped_by_employee[record.employee_id].append(record)

                fl = self.add_employee_attendance_data(grouped_by_employee, report_date_start_from, report_date_end_to)

            output = base64.encodebytes(fl[1])

            context = self.env.args
            ctx = dict(context[2])
            ctx.update({'employee_attendance_report_file_store': output, 'file': fl[0]})

            self.employee_attendance_report_name = fl[0]
            self.employee_attendance_report_file_store = ctx['employee_attendance_report_file_store']
            self.is_printed = True

            return {
                'type': 'ir.actions.act_window',
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'employee.attendance.report',
                'target': 'new',
                'context': ctx,
                'res_id': self.id,
            }

        except Exception as e:
            raise UserError(_("An error occurred while generating the attendance report. "
                              "Please try again or contact the administrator. Details: %s") % str(e))

    def action_go_backword(self):
        """
        Reset Report View and Navigate Back.

        Purpose:
        --------
        This method is typically triggered by a 'Back' button in the Attendance Report wizard.
        It resets the state of the wizard (marks `is_printed` as False) and returns an
        action to reload the same wizard form view.

        Workflow:
        ---------
        1. Ensure the context exists; if not, initialize it as an empty dictionary.
        2. Reset `is_printed` flag to False to indicate the report is not currently shown.
        3. Return an Odoo action dictionary that reloads the current wizard form in a new window.

        Returns:
        --------
        dict:
            A standard Odoo `ir.actions.act_window` dictionary to display the wizard form again.

        Exceptions:
        -----------
        Raises a professional `UserError` if any unexpected error occurs during execution.
        """
        try:
            if self._context is None:
                self._context = {}

            # Reset the printed status
            self.is_printed = False

            back_action = {
                'type': 'ir.actions.act_window',
                'view_type': 'form',
                'view_mode': 'form',
                'res_model': 'employee.attendance.report',
                'target': 'new',
            }

            return back_action

        except Exception as error:
            raise UserError(
                _("Unable to navigate back to the report view. Please try again or contact support. "
                  "Details: %s") % str(error)
            )
    def add_employee_attendance_data(self, attendances, report_date_start_from, report_date_end_to):

        str_date1 = str(self.report_date_start_from)
        str_date1 = self.new_timezone(self.report_date_start_from)

        date1 = datetime.strptime(str_date1, '%Y-%m-%d %H:%M:%S').date()
        day1 = date1.strftime('%d')
        month1 = date1.strftime('%B')
        year1 = date1.strftime('%Y')
        str_date2 = str(self.report_date_end_to)
        str_date2 = self.new_timezone(self.report_date_end_to)
        date2 = datetime.strptime(str_date2, '%Y-%m-%d %H:%M:%S').date()
        day2 = date2.strftime('%d')
        month2 = date2.strftime('%B')
        year2 = date2.strftime('%Y')
        fl = 'Attendance from ' + day1 + '-' + month1 + '-' + year1 + ' to ' + day2 + '-' + month2 + '-' + year2 + '(' + str(
            datetime.today()) + ')' + '.xlsx'
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output)
        param_val = self.env['ir.config_parameter'].sudo().get_param(
            'tis_hr_biometric_attendance.multiple_shift')
        if param_val in [False, 'False', 'false', '0', 0,None, '']:  # Not True
            for employee, records in attendances.items():
                # for attendance in records:
                if self.attendance_excel_sheet_name:
                    if self.attendance_excel_sheet_name == 'badge_id' and employee.barcode:
                        worksheet = workbook.add_worksheet(f'{employee.barcode}')
                    else:
                        worksheet = workbook.add_worksheet(f'{employee.name}')
                else:
                    worksheet = workbook.add_worksheet(f'{employee.name}')

                # worksheet = workbook.add_worksheet(f'{employee.name}')
                worksheet.set_landscape()

                bold = workbook.add_format({'bold': True, 'border': 1,
                                            'align': 'center',
                                            'font_size': 15})
                emp_name_style_en = workbook.add_format({'bold': True, 'border': 1,
                                            'align': 'left',
                                            'font_size': 13})
                emp_name_style_ar = workbook.add_format({'bold': True, 'border': 1,
                                                         'align': 'right',
                                                         'font_size': 13})

                font_left = workbook.add_format({'align': 'left',
                                                 'border': 1,
                                                 'font_size': 12})
                font_center = workbook.add_format({'align': 'center',
                                                   'border': 1,
                                                   'valign': 'vcenter',
                                                   'font_size': 12})
                font_center_o = workbook.add_format({'align': 'center',
                                                   'border': 1,
                                                   'valign': 'vcenter',
                                                   'font_size': 12})
                font_bold_center = workbook.add_format({'align': 'center',
                                                        'border': 1,
                                                        'valign': 'vcenter',
                                                        'font_size': 12,
                                                        'bold': True})
                font_bold_left = workbook.add_format({'align': 'left',
                                                        'border': 1,
                                                        'valign': 'vcenter',
                                                        'font_size': 12,
                                                        'bold': True})
                font_left_ab = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'font_size': 12,
                    'bg_color': '#FFFF00'  # Yellow background
                })

                font_center_ab = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00'  # Yellow background
                })

                absent_weeend_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',  # Yellow background
                    'font_color': '#ad1111'
                })
                medical_leave_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#2ac5f5',  # sky blue background
                    'font_color': '#050505'
                })
                vacation_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#23cc2e',  # Green background
                    'font_color': '#fafafa'
                })
                vacation_working_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#23cc2e',  # Green background
                    'font_color': '#fafafa' # white text
                })
                holiday_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#f79431',  # Orange background
                    'font_color': '#fafafa'
                })

                holiday_present_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#f79431',  # Orange background
                    'font_color': '#fafafa' # white text
                })

                present_weeend_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',  # Yellow background
                    'font_color': '#006400'  # Dark green font color
                })

                present_name_weeend_style = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',  # Yellow background
                    'font_color': '#006400'  # Dark green font color
                })

                working_day_absent_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#edbebe',
                    'font_color': '#ad1111'  # Dark green font color
                })
                name_working_day_absent_style = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#edbebe',
                    'font_color': '#ad1111'  # Dark green font color
                })
                name_weekend_absent_style = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',
                    'font_color': '#ad1111'  # Dark green font color
                })

                border = workbook.add_format({'border': 1})

                # worksheet.set_column('0:XFD', None, None, {'hidden': True})
                worksheet.set_column('A:O', 20, border)
                worksheet.set_row(0, 20)
                worksheet.merge_range('A1:J1',
                                      "Attendance sheet from" + day1 + '-' + month1 + '-' + year1 + ' to ' + day2 + '-' + month2 + '-' + year2,
                                      bold)
                # employee name hin english
                worksheet.merge_range('A2:F2',employee.name or ' ' ,emp_name_style_en)
                worksheet.merge_range('G2:K2', employee.arabic_name or ' ', emp_name_style_ar)

                row = 3
                col = 0
                worksheet.merge_range(row, col + 0, row + 1, col + 0, "SR No.", font_bold_center)
                worksheet.merge_range(row, col + 1, row + 1, col + 1, "Day", font_bold_center)
                worksheet.merge_range(row, col + 2, row + 1, col + 2, "Date", font_bold_center)
                worksheet.merge_range(row, col + 3, row + 1, col + 3, "Check In", font_bold_center)
                worksheet.merge_range(row, col + 4, row + 1, col + 4, "Check_out", font_bold_center)
                worksheet.merge_range(row, col + 5, row + 1, col + 5, "Difference", font_bold_center)
                worksheet.merge_range(row, col + 6, row + 1, col + 6, "Break Time", font_bold_center)
                worksheet.merge_range(row, col + 7, row + 1, col + 7, "Worked Hours", font_bold_center)
                worksheet.merge_range(row, col + 8, row + 1, col + 8, "Shift Hours", font_bold_center)
                worksheet.merge_range(row, col + 9, row + 1, col + 9, "Overtime Hours", font_bold_center)
                worksheet.merge_range(row, col + 10, row + 1, col + 10, "Shortfall Hours", font_bold_center)
                row += 2
                total_overtime = "00:00"
                total_overtime_hours_rounded = 0
                total_overtime_minutes_rounded = 0
                total_ot_list = []
                total_working_hours = 0
                total_shortfall_hours = timedelta()
                total_worked_hours = timedelta()
                # my_dates = datetime.strptime(date_from, "%Y-%m-%d %H:%M:%S")  # Convert to datetime
                my_date = datetime.strptime(report_date_start_from, '%Y-%m-%d %H:%M:%S').date()  # Convert date_from to date
                last_date = datetime.strptime(report_date_end_to, '%Y-%m-%d %H:%M:%S').date()
                last_date += timedelta(days=1)
                # date_to_dt = datetime.strptime(date_to, "%Y-%m-%d %H:%M:%S")  # Convert to datetime
                # my_date = my_dates
                records = sorted(records, key=lambda r: r.check_in)
                total_absent_hours = 0
                total_day_in_months = 0
                total_friday_in_month = 0
                total_absent_in_month =  0
                sr_no_count = 1
                for attendance in records:

                    check_in_date = attendance.check_in.date() if attendance.check_in else None
                    check_out_date = attendance.check_out.date() if attendance.check_out else None
                    date_check = check_in_date if check_in_date else None
                    count = 0
                    while date_check and my_date < date_check:  # Run only until the last recorded date
                        is_absent = False

                        if attendance.check_in and attendance.check_out:
                            if check_in_date != my_date:
                                is_absent = True  # Mark as absent if check_in doesn't match

                        elif attendance.check_in and not attendance.check_out:
                            if check_in_date != my_date:
                                is_absent = True  # Mark as absent if check_in doesn't match

                        if is_absent:
                            font_to_set = working_day_absent_style
                            font_to_set_name = name_working_day_absent_style
                            week_of = my_date and my_date.weekday() == 4
                            if week_of:
                                font_to_set = absent_weeend_style
                                font_to_set_name = name_weekend_absent_style
                            if is_absent and not week_of:
                                total_absent_in_month += 1
                                total_day_in_months += 1
                            elif is_absent and  week_of:
                                total_friday_in_month += 1
                                total_day_in_months += 1

                            # Get default font
                            ef = font_to_set

                            # Initialize leave tracking
                            leave_type = None
                            leave_style = None
                            font_to_set = None

                            ml_chek_in = False
                            ml_chek_out = False
                            working_hours = False
                            ml_difference = False

                            # Determine leave type on the given date
                            for leave in employee.leave_line_ids:
                                if leave.date == my_date:
                                    leave_type = leave.leave_type
                                    if leave_type == 'medical':
                                        if leave.att_start_date and leave.att_end_date:
                                            ml_chek_in = leave.att_start_date.strftime("%Y-%m-%d %H:%M:%S")
                                            ml_chek_out = leave.att_end_date.strftime("%Y-%m-%d %H:%M:%S")
                                            duration = leave.att_end_date - leave.att_start_date
                                            wh = duration.total_seconds() / 3600 if ml_chek_in and ml_chek_out else 0.0
                                            if wh > 0:
                                                working_hours = wh
                                                ml_difference = working_hours
                                    break

                            # Map leave type to style
                            if leave_type == 'holiday':
                                leave_style = holiday_style
                                status_label = 'Holiday'
                            elif leave_type == 'vacation':
                                leave_style = vacation_style
                                status_label = 'Vacation'
                            elif leave_type == 'medical':
                                leave_style = medical_leave_style
                                status_label = 'Medical Leave'
                            else:
                                status_label = 'Weekend' if week_of else 'Absent'

                            # Determine font to apply
                            if leave_type and not week_of:
                                font_to_set = leave_style
                            else:
                                font_to_set = ef

                            # Write values to Excel
                            worksheet.write(row, col + 0, sr_no_count, font_to_set)
                            sr_no_count += 1


                            worksheet.write(row, col + 1, my_date.strftime("%A"), font_to_set)
                            worksheet.write(row, col + 2, my_date.strftime("%Y-%m-%d"), font_to_set)
                            worksheet.write(row, col + 3, ml_chek_in or status_label, font_to_set)
                            worksheet.write(row, col + 4, ml_chek_out or status_label, font_to_set)
                            worksheet.write(row, col + 5, float(ml_difference) if ml_difference else ' ', font_to_set)
                            worksheet.write(row, col + 6, ' ', font_to_set)
                            worksheet.write(row, col + 7, float(working_hours) if working_hours else ' ', font_to_set)

                            # Set working hours if not on leave or weekend
                            if leave_type in ['holiday', 'vacation', 'medical']:
                                worksheet.write(row, col + 8, '', font_to_set)
                                worksheet.write(row, col + 10, '', font_to_set)
                            else:
                                worksheet.write(row, col + 8,
                                                -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                                font_to_set)
                                worksheet.write(row, col + 10,
                                                -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                                font_to_set)

                            worksheet.write(row, col + 9, ' ', font_to_set)

                            if is_absent and not week_of:
                                total_absent_hours += employee.resource_calendar_id.working_hours or 8
                            row += 1
                            my_date += timedelta(days=1)
                        # ✅ Move to the next date to prevent infinite loop
                    p_on_weekend = True if my_date.weekday() == 4 else False
                    # worksheet.merge_range(row, col, row, col + 1, attendance.employee_id.name, present_name_weeend_style if p_on_weekend else font_left)
                    #####
                    ef = font_center
                    if str(my_date) == '2025-06-11':
                        print(font_center)
                    # Initialize leave tracking
                    leave_type = None
                    leave_style = None
                    font_to_set = None

                    # Determine leave type on the given date
                    for leave in employee.leave_line_ids:
                        if leave.date == my_date:
                            leave_type = leave.leave_type
                            break

                    # Map leave type to style
                    if leave_type == 'holiday':
                        leave_style = holiday_present_style
                    elif leave_type == 'vacation':
                        leave_style = vacation_working_style
                    elif leave_type == 'medical':
                        leave_style = medical_leave_style

                    # Determine font to apply
                    # font_center = leave_style if leave_type and not p_on_weekend else ef


                    leave_ot = False
                    if leave_type == 'holiday' or leave_type == 'vacation':
                        leave_ot = True

                    ##### #### #####
                    ##### #### #####
                    ##### #### #####

                    worksheet.write(row, col + 0, sr_no_count, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    sr_no_count += 1

                    if my_date.weekday() == 4:
                        total_friday_in_month += 1
                        total_day_in_months += 1
                    elif my_date.weekday() != 4:
                        total_day_in_months += 1
                    if attendance.check_in:
                        check_in = self.new_timezone(attendance.check_in)
                        if isinstance(check_in, str):  # Convert string to datetime
                            check_in = datetime.fromisoformat(check_in)
                        day_name = check_in.strftime("%A")
                        date_name = check_in.strftime("%Y-%m-%d")
                    elif attendance.check_out:
                        check_out = self.new_timezone(attendance.check_out)
                        if isinstance(check_out, str):  # Convert string to datetime
                            check_out = datetime.fromisoformat(check_out)
                        day_name = check_out.strftime("%A")
                        date_name = check_out.strftime("%Y-%m-%d")
                    else:
                        day_name = "N/A"
                        date_name = "N/A"
                    worksheet.write(row, col + 1, day_name, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    worksheet.write(row, col + 2, date_name, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    if attendance.check_in:
                        check_in = self.new_timezone(attendance.check_in)
                    else:
                        check_in = '***No Check In***'
                    worksheet.write(row, col + 3, check_in, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    if attendance.check_out:
                        check_out = self.new_timezone(attendance.check_out)
                    else:
                        check_out = '***No Check Out***'

                    worksheet.write(row, col + 4, check_out, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    diff_hours = int(attendance.check_in_check_out_difference)
                    diff_minutes = int((attendance.check_in_check_out_difference - diff_hours) * 60)
                    diff = f"{diff_hours:02}:{diff_minutes:02}"
                    worksheet.write(row, col + 5, diff, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    new_worked_hours = attendance.check_in_check_out_difference - attendance.break_time
                    wrk_hours = int(new_worked_hours)
                    wrk_minutes = int((new_worked_hours - wrk_hours) * 60)
                    # wrk_hours = int(attendance.worked_hours)
                    # wrk_minutes = int((attendance.worked_hours - wrk_hours) * 60)
                    worked = f"{wrk_hours:02}:{wrk_minutes:02}"
                    if worked:
                        total_worked_hours += timedelta(hours=wrk_hours, minutes=wrk_minutes)
                    worksheet.write(row, col + 7, worked, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    def parse_time_to_minutes(t):
                        sign = -1 if t.startswith('-') else 1
                        t = t.lstrip('-')
                        hours, minutes = map(int, t.split(":"))
                        return sign * (hours * 60 + minutes)

                    # Replace this block where break_time is calculated
                    if diff.startswith('-') or worked.startswith('-'):
                        diff_minutes = parse_time_to_minutes(diff)
                        worked_minutes = parse_time_to_minutes(worked)
                        break_time_minutes = diff_minutes - worked_minutes
                        break_time = timedelta(minutes=break_time_minutes)
                    else:
                        break_time = datetime.strptime(diff, "%H:%M") - datetime.strptime(worked, "%H:%M")

                    a_new_break_time = self.calculate_break_time(attendance)
                    def convert_decimal_to_time(decimal_hours):
                        hours = int(decimal_hours)
                        minutes = round((decimal_hours - hours) * 60)
                        return f"{hours:02d}:{minutes:02d}"  # Format as HH:MM
                    if not a_new_break_time:
                        new_break_time = convert_decimal_to_time(attendance.break_time) if attendance.break_time else "00:00"
                    else:
                        new_break_time =  a_new_break_time if a_new_break_time else "00:00"

                    hours, remainder = divmod(total_worked_hours.seconds, 3600)
                    minutes = remainder // 60
                    formatted_break_time = f"{hours:02}:{minutes:02}"
                    worksheet.write(row, col + 6, new_break_time if new_break_time else formatted_break_time, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    # Determine if the current day is Friday
                    check_in_date = datetime.strptime(check_in, "%Y-%m-%d %H:%M:%S") if attendance.check_in else None
                    is_friday = check_in_date and check_in_date.weekday() == 4  # 4 corresponds to Friday

                    # Write the shift hours based on whether it is Friday
                    if is_friday or leave_ot:
                        worksheet.write(row, col + 8, "00:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    elif employee.resource_calendar_id and employee.resource_calendar_id.working_hours:
                        worksheet.write(row, col + 8, employee.resource_calendar_id.working_hours,
                                        present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        total_working_hours += employee.resource_calendar_id.working_hours

                    else:
                        worksheet.write(row, col + 8, "08:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        total_working_hours += 8

                    def calculate_time(input_time, check_in, employee, base_hours=8):
                        cumulative_overtime = timedelta(0)
                        cumulative_shortfall = timedelta(0)

                        # Convert input_time to timedelta
                        hours, minutes = map(int, input_time.split(":"))
                        input_timedelta = timedelta(hours=hours, minutes=minutes)
                        base_hours = 8
                        if employee and employee.resource_calendar_id and employee.resource_calendar_id.working_hours:
                            if int(employee.resource_calendar_id.working_hours) > 0:
                                base_hours = int(employee.resource_calendar_id.working_hours)

                        base_time = timedelta(hours=base_hours)

                        # Convert check_in to a datetime object to determine the day of the week
                        check_in_date = datetime.strptime(check_in, "%Y-%m-%d %H:%M:%S")
                        is_friday = check_in_date.weekday() == 4  # 4 corresponds to Friday

                        # Calculate adjusted time
                        adjusted_timedelta = input_timedelta - base_time

                        leave_type = None


                        # Determine leave type on the given date
                        for leave in employee.leave_line_ids:
                            if leave.date == my_date:
                                leave_type = leave.leave_type
                                break

                        leave_ot = False
                        if leave_type == 'holiday' or leave_type == 'vacation':
                            leave_ot = True

                        if is_friday:
                            cumulative_overtime = input_time
                            classification = "Overtime"
                            return cumulative_overtime, classification
                        elif not is_friday and leave_ot:
                            cumulative_overtime = input_time
                            classification = "Overtime"
                            return cumulative_overtime, classification
                        else:
                            # Update cumulative counters and classify
                            if adjusted_timedelta > timedelta(0) or is_friday:
                                cumulative_overtime += adjusted_timedelta
                                classification = "Overtime"
                            else:
                                cumulative_shortfall -= adjusted_timedelta  # Subtract because adjusted_timedelta is negative
                                classification = "Shortfall"

                            # Extract hours and minutes
                            total_seconds = adjusted_timedelta.total_seconds()
                            total_minutes = total_seconds // 60
                            overtime_hours = total_minutes // 60
                            overtime_minutes = total_minutes % 60

                            # Apply corrected overtime rounding logic based on image
                            if overtime_hours == 0:
                                if 0 <= overtime_minutes <= 29:
                                    overtime_minutes = 0
                                elif 30 <= overtime_minutes <= 44:
                                    overtime_minutes = 30
                                elif 45 <= overtime_minutes <= 50:
                                    overtime_minutes = 45
                                elif 51 <= overtime_minutes <= 59:
                                    overtime_minutes = 0
                                    overtime_hours = 1
                            else:
                                if 0 <= overtime_minutes <= 29:
                                    overtime_minutes = 0
                                elif 30 <= overtime_minutes <= 44:
                                    overtime_minutes = 30
                                elif 45 <= overtime_minutes <= 50:
                                    overtime_minutes = 45
                                elif 51 <= overtime_minutes <= 59:
                                    overtime_minutes = 0
                                    overtime_hours += 1

                            formatted_adjusted_time = f"{int(overtime_hours):02}:{int(overtime_minutes):02}"

                            return formatted_adjusted_time, classification

                    # Function to format cumulative time for display
                    def format_cumulative_time(timedelta_obj):
                        hours, remainder = divmod(timedelta_obj.seconds, 3600)
                        minutes = remainder // 60
                        return f"{hours:02}:{minutes:02}"

                    # Example usage:
                    adjusted_time, classification = calculate_time(worked, check_in, employee)  # Friday

                    def round_time(adjusted_time):
                        hours, minutes = map(int, adjusted_time.split(":"))
                        total_minutes = hours * 60 + minutes
                        new_hours = total_minutes // 60
                        remaining_minutes = total_minutes % 60
                        if new_hours <= 0:
                            if 0 <= remaining_minutes <= 29:
                                new_minutes = 0
                            elif 30 <= remaining_minutes <= 44:
                                new_minutes = 30
                            elif 45 <= remaining_minutes <= 50:
                                new_minutes = 45
                            else:  # 51 to 59
                                new_hours += 1
                                new_minutes = 0
                        elif new_hours > 0:
                            if 0 <= remaining_minutes <= 14:
                                new_minutes = 0
                            elif 15 <= remaining_minutes <= 45:
                                new_minutes = 30
                            elif 46 <= remaining_minutes <= 59:
                                new_minutes = 0
                                new_hours += 1  # Round up to the next hour

                        return f"{new_hours:02}:{new_minutes:02}"
                    if adjusted_time and my_date.weekday() == 4:

                        rounded_time = round_time(adjusted_time)
                        # time_spited = rounded_time.split(':')
                        # total_overtime_hours_rounded += float(time_spited[0]) or 0
                        # total_overtime_minutes_rounded += float(time_spited[1]) or 0
                        total_ot_list.append(rounded_time)
                        worksheet.write(row, col + 9, rounded_time if classification == "Overtime" else "00:00",
                                        present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    else:
                        rounded_time = round_time(adjusted_time)
                        # time_spited = rounded_time.split(':')
                        # total_overtime_hours_rounded += float(time_spited[0]) or 0
                        # total_overtime_minutes_rounded += float(time_spited[1]) or 0
                        total_ot_list.append(rounded_time)
                        worksheet.write(row, col + 9, rounded_time if classification == "Overtime" else "00:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    hours1, minutes1 = map(int, total_overtime.split(":"))
                    hours2, minutes2 = map(int,
                                           adjusted_time.split(":") if classification == "Overtime" else "00:00".split(":"))

                    time1_delta = timedelta(hours=hours1, minutes=minutes1)
                    time2_delta = timedelta(hours=hours2, minutes=minutes2)
                    # Add the two timedelta objects
                    total_time = time1_delta + time2_delta

                    # Convert back to "HH:MM" format
                    total_hours, remainder = divmod(total_time.seconds, 3600)
                    total_minutes = remainder // 60
                    total_overtime = f"{total_time.days * 24 + total_hours:02}:{total_minutes:02}"



                    worksheet.write(row, col + 10, adjusted_time if classification == "Shortfall" else "00:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    if classification == "Shortfall" and adjusted_time:
                        hours, minutes = map(int, adjusted_time.split(":"))
                        total_shortfall_hours += timedelta(hours=hours, minutes=minutes)

                    my_date += timedelta(days=1)
                    row += 1
                if my_date < last_date:
                    while my_date < last_date:
                        font_to_set = working_day_absent_style
                        font_to_set_name = name_working_day_absent_style
                        week_of = my_date and my_date.weekday() == 4
                        if week_of:
                            font_to_set = absent_weeend_style
                            font_to_set_name = name_weekend_absent_style
                            total_friday_in_month += 1
                            total_day_in_months += 1
                        else:
                            total_day_in_months += 1
                            total_absent_in_month += 1

                        ###
                        ef = font_to_set
                        leave_type = None
                        leave_style = None
                        font_to_set = None

                        ml_chek_in = False
                        ml_chek_out = False
                        working_hours = False
                        ml_difference = False


                        # Determine leave type on the given date
                        for leave in employee.leave_line_ids:
                            if leave.date == my_date:
                                leave_type = leave.leave_type
                                if leave_type == 'medical':
                                    if leave.att_start_date and leave.att_end_date:
                                        ml_chek_in = leave.att_start_date.strftime("%Y-%m-%d %H:%M:%S")
                                        ml_chek_out = leave.att_end_date.strftime("%Y-%m-%d %H:%M:%S")
                                        duration = leave.att_end_date - leave.att_start_date
                                        wh = duration.total_seconds() / 3600 if ml_chek_in and ml_chek_out else 0.0
                                        if wh > 0:
                                            working_hours = wh
                                            ml_difference = working_hours
                                break

                        # Apply style if present and check-in exists
                        if leave_type and attendance and attendance.check_in:
                            if leave_type == 'holiday':
                                leave_style = holiday_style
                            elif leave_type == 'vacation':
                                leave_style = vacation_style
                            elif leave_type == 'medical':
                                leave_style = medical_leave_style

                        # Apply font if it's a leave day and not weekend
                        if leave_type and not week_of:
                            font_to_set = leave_style
                        else:
                            font_to_set = ef

                        # Set label depending on leave or weekend
                        if week_of:
                            status_label = 'Weekend'
                        elif leave_type == 'holiday':
                            status_label = 'Holiday'
                        elif leave_type == 'vacation':
                            status_label = 'Vacation'
                        elif leave_type == 'medical':
                            status_label = 'Medical Leave'
                        else:
                            status_label = 'Absent'

                        # worksheet.merge_range(row, col, row, col + 1, attendance.employee_id.name, font_to_set_name)
                        worksheet.write(row, col + 0,sr_no_count, font_to_set)
                        sr_no_count += 1

                        # Write values to Excel
                        worksheet.write(row, col + 1, my_date.strftime("%A"), font_to_set)
                        worksheet.write(row, col + 2, my_date.strftime("%Y-%m-%d"), font_to_set)
                        worksheet.write(row, col + 3, ml_chek_in or status_label, font_to_set)
                        worksheet.write(row, col + 4, ml_chek_out or status_label, font_to_set)
                        worksheet.write(row, col + 5, float(ml_difference) or ' ', font_to_set)
                        worksheet.write(row, col + 6, ' ', font_to_set)
                        worksheet.write(row, col + 7, float(working_hours) or ' ', font_to_set)

                        # If leave_type in holiday/vacation/medical, set '' in hours, else working hours
                        if leave_type in ['holiday', 'vacation', 'medical']:
                            worksheet.write(row, col + 8, '', font_to_set)
                            worksheet.write(row, col + 10, '', font_to_set)
                        else:
                            worksheet.write(row, col + 8,
                                            -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                            font_to_set)
                            worksheet.write(row, col + 10,
                                            -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                            font_to_set)

                        worksheet.write(row, col + 9, ' ', font_to_set)

                        if not week_of:
                            total_absent_hours += employee.resource_calendar_id.working_hours or 8
                        my_date += timedelta(days=1)
                        row += 1
                ####
                new_over_time = ''
                # Convert '91:57' to total minutes
                total_overtime_minutes = sum(
                    (int(ots.split(':')[0]) * 60 + int(ots.split(':')[1]))
                    if ots and ':' in ots and '-' not in ots.split(':')[0] else 0
                    for ots in total_ot_list
                )

                # hours, minutes = map(int, "91:57".split(':'))
                # total_overtime_minutes = (total_overtime_hours_rounded * 60) + total_overtime_minutes_rounded
                # Convert 112 hours to minutes
                total_absent_minutes = total_absent_hours * 60  # 112*60 = 6720 minutes


                # Subtract minutes
                # remaining_minutes = total_overtime_minutes - total_absent_minutes
                remaining_minutes = total_overtime_minutes


                # Convert back to hours and minutes
                remaining_hours = remaining_minutes // 60
                remaining_mins = abs(remaining_minutes % 60)  # Use abs() to avoid negative minutes
                new_over_time = f"{remaining_hours}:{remaining_mins:02d}"
                # Format result
                # if remaining_minutes < 0:
                #     new_over_time = f"-{abs(remaining_hours)}:{remaining_mins:02d}"  # Handle negative time
                # else:
                #     new_over_time = f"{remaining_hours}:{remaining_mins:02d}"

                font_center = font_center_o
                worksheet.write(row, 9, f"Total: {new_over_time}", font_center)
                # worksheet.write(row, 10, f"Total: {total_overtime}", font_center)
                worksheet.write(row, 8, f"Total: {total_working_hours}", font_center)
                # Convert timedelta to total hours and minutes
                total_hours = total_shortfall_hours.seconds // 3600  # Extract hours
                total_minutes = (total_shortfall_hours.seconds % 3600) // 60  # Extract minutes
                #####
                total_seconds = total_worked_hours.total_seconds()
                total_hours = int(total_seconds // 3600)
                total_minutes = int((total_seconds % 3600) // 60)
                formatted_worked_time = f"{total_hours:02}:{total_minutes:02}"
                # worksheet.write(row, 11, f"Total: {total_hours:02}:{total_minutes:02}", font_center)
                worksheet.write(row, 7, f"Total: {formatted_worked_time}", font_center)

                #calculation for net overtime
                net_total_overtime_mins = total_overtime_minutes - total_absent_minutes
                nt_total_seconds = net_total_overtime_mins * 60
                nt_total_hours = int(nt_total_seconds // 3600)
                nt_total_minutes = int((nt_total_seconds % 3600) // 60)
                nt_formatted_ot_time = f"{nt_total_hours:02}:{nt_total_minutes:02}"


                # totals
                worksheet.merge_range(row + 2, 0, row + 2, 1, 'Total Days in month', font_bold_left)
                worksheet.merge_range(row + 3, 0, row + 3, 1, 'Total Friday', font_bold_left)
                worksheet.merge_range(row + 4, 0, row + 4, 1, 'Working days in a month', font_bold_left)
                worksheet.merge_range(row + 5, 0, row + 5, 1, 'Absent / Shortfall', font_bold_left)
                worksheet.merge_range(row + 6, 0, row + 6, 1, 'Total Absent Hours', font_bold_left)
                worksheet.merge_range(row + 7, 0, row + 7, 1, 'Total Overtime on month', font_bold_left)
                worksheet.merge_range(row + 8, 0, row + 8, 1, 'Net Overtime', font_bold_left)

                # worksheet.merge_range(row, col, row, col + 1, attendance.employee_id.name, font_to_set_name)

                worksheet.write(row + 2, 2, total_day_in_months, font_bold_center)
                worksheet.write(row + 3, 2, total_friday_in_month, font_bold_center)
                worksheet.write(row + 4, 2, total_day_in_months - total_friday_in_month, font_bold_center)
                worksheet.write(row + 5, 2, total_absent_in_month, font_bold_center)
                worksheet.write(row + 6, 2, total_absent_hours, font_bold_center)
                worksheet.write(row + 7, 2, new_over_time if new_over_time else '', font_bold_center)
                worksheet.write(row + 8, 2, nt_formatted_ot_time or '', font_bold_center)

        else:
            #
            for employee, records in attendances.items():
                if self.attendance_excel_sheet_name:
                    if self.attendance_excel_sheet_name == 'badge_id' and employee.barcode:
                        worksheet = workbook.add_worksheet(f'{employee.barcode}')
                    else:
                        worksheet = workbook.add_worksheet(f'{employee.name}')
                else:
                    worksheet = workbook.add_worksheet(f'{employee.name}')

                # worksheet = workbook.add_worksheet(f'{employee.name}')
                worksheet.set_landscape()

                bold = workbook.add_format({'bold': True, 'border': 1,
                                            'align': 'center',
                                            'font_size': 15})
                emp_name_style_en = workbook.add_format({'bold': True, 'border': 1,
                                                         'align': 'left',
                                                         'font_size': 13})
                emp_name_style_ar = workbook.add_format({'bold': True, 'border': 1,
                                                         'align': 'right',
                                                         'font_size': 13})

                font_left = workbook.add_format({'align': 'left',
                                                 'border': 1,
                                                 'font_size': 12})
                font_center = workbook.add_format({'align': 'center',
                                                   'border': 1,
                                                   'valign': 'vcenter',
                                                   'font_size': 12})
                font_bold_center = workbook.add_format({'align': 'center',
                                                        'border': 1,
                                                        'valign': 'vcenter',
                                                        'font_size': 12,
                                                        'bold': True})
                font_bold_left = workbook.add_format({'align': 'left',
                                                      'border': 1,
                                                      'valign': 'vcenter',
                                                      'font_size': 12,
                                                      'bold': True})
                font_left_ab = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'font_size': 12,
                    'bg_color': '#FFFF00'  # Yellow background
                })

                font_center_ab = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00'  # Yellow background
                })

                absent_weeend_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',  # Yellow background
                    'font_color': '#ad1111'
                })

                present_weeend_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',  # Yellow background
                    'font_color': '#006400'  # Dark green font color
                })

                present_name_weeend_style = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',  # Yellow background
                    'font_color': '#006400'  # Dark green font color
                })

                medical_leave_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#2ac5f5',  # sky blue background
                    'font_color': '#050505'
                })
                vacation_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#23cc2e',  # Green background
                    'font_color': '#fafafa'
                })
                vacation_working_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#23cc2e',  # Green background
                    'font_color': '#fafafa'  # white text
                })
                holiday_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#f79431',  # Orange background
                    'font_color': '#fafafa'
                })

                holiday_present_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#f79431',  # Orange background
                    'font_color': '#fafafa'  # white text
                })

                working_day_absent_style = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#edbebe',
                    'font_color': '#ad1111'  # Dark green font color
                })
                name_working_day_absent_style = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#edbebe',
                    'font_color': '#ad1111'  # Dark green font color
                })
                name_weekend_absent_style = workbook.add_format({
                    'align': 'left',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#FFFF00',
                    'font_color': '#ad1111'  # Dark green font color
                })

                multiple_shift_font = workbook.add_format({
                    'align': 'center',
                    'border': 1,
                    'valign': 'vcenter',
                    'font_size': 12,
                    'bg_color': '#e6ffcc',
                    'font_color': '#141414'
                })

                border = workbook.add_format({'border': 1})

                # worksheet.set_column('0:XFD', None, None, {'hidden': True})
                worksheet.set_column('A:O', 20, border)
                worksheet.set_row(0, 20)
                worksheet.merge_range('A1:L1',
                                      "Attendance Sheet From " + "(" + day1 + '-' + month1 + '-' + year1 + ' to ' + day2 + '-' + month2 + '-' + year2 + ")",
                                      bold)
                # employee name hin english
                worksheet.merge_range('A2:F2', employee.name or ' ', emp_name_style_en)
                worksheet.merge_range('G2:L2', employee.arabic_name or ' ', emp_name_style_ar)

                row = 3
                col = 0
                worksheet.merge_range(row, col + 0, row + 1, col + 0, "SR No.", font_bold_center)
                worksheet.merge_range(row, col + 1, row + 1, col + 1, "Day", font_bold_center)
                worksheet.merge_range(row, col + 2, row + 1, col + 2, "Date", font_bold_center)
                worksheet.merge_range(row, col + 3, row + 1, col + 3, "Check In", font_bold_center)
                worksheet.merge_range(row, col + 4, row + 1, col + 4, "Check_out", font_bold_center)
                worksheet.merge_range(row, col + 5, row + 1, col + 5, "Difference", font_bold_center)
                worksheet.merge_range(row, col + 6, row + 1, col + 6, "Break Time", font_bold_center)
                worksheet.merge_range(row, col + 7, row + 1, col + 7, "Worked Hours", font_bold_center)
                worksheet.merge_range(row, col + 8, row + 1, col + 8, "Total Worked Hours", font_bold_center)

                worksheet.merge_range(row, col + 9, row + 1, col + 9, "Shift Hours", font_bold_center)
                worksheet.merge_range(row, col + 10, row + 1, col + 10, "Overtime Hours", font_bold_center)
                worksheet.merge_range(row, col + 11, row + 1, col + 11, "Shortfall Hours", font_bold_center)
                row += 2
                total_overtime = "00:00"
                total_overtime_hours_rounded = 0
                total_overtime_minutes_rounded = 0
                total_ot_list = []
                total_working_hours = 0
                total_shortfall_hours = timedelta()
                total_worked_hours = timedelta()
                # my_dates = datetime.strptime(date_from, "%Y-%m-%d %H:%M:%S")  # Convert to datetime
                my_date = datetime.strptime(report_date_start_from, '%Y-%m-%d %H:%M:%S').date()  # Convert date_from to date
                last_date = datetime.strptime(report_date_end_to, '%Y-%m-%d %H:%M:%S').date()
                last_date += timedelta(days=1)
                # date_to_dt = datetime.strptime(date_to, "%Y-%m-%d %H:%M:%S")  # Convert to datetime
                # my_date = my_dates
                records = sorted(records, key=lambda r: r.check_in)
                total_absent_hours = 0
                total_day_in_months = 0
                total_friday_in_month = 0
                total_absent_in_month = 0
                sr_no_count = 1
                for attendance in records:
                    leave_type = attendance.leave_type
                    multi_punching_count = 0
                    multi_punching_records = True  if len(attendance.multiple_checkin_ids) > 1 else False
                    if multi_punching_records:
                        multi_punching_count = len(attendance.multiple_checkin_ids)

                    check_in_date = attendance.check_in.date() if attendance.check_in else None
                    check_out_date = attendance.check_out.date() if attendance.check_out else None
                    date_check = check_in_date if check_in_date else None

                    while date_check and my_date != date_check:  # Run only until the last recorded date
                        is_absent = False

                        if attendance.check_in and attendance.check_out:
                            if check_in_date != my_date:
                                is_absent = True  # Mark as absent if check_in doesn't match

                        elif attendance.check_in and not attendance.check_out:
                            if check_in_date != my_date:
                                is_absent = True  # Mark as absent if check_in doesn't match

                        if is_absent:
                            font_to_set = working_day_absent_style
                            font_to_set_name = name_working_day_absent_style
                            week_of = my_date and my_date.weekday() == 4
                            if week_of:
                                font_to_set = absent_weeend_style
                                font_to_set_name = name_weekend_absent_style
                            if is_absent and not week_of:
                                total_absent_in_month += 1
                                total_day_in_months += 1
                            elif is_absent and week_of:
                                total_friday_in_month += 1
                                total_day_in_months += 1

                            ####
                            # Get default font
                            ef = font_to_set

                            # Initialize leave tracking
                            leave_type = None
                            leave_style = None
                            font_to_set = None

                            # Determine leave type on the given date
                            for leave in employee.leave_line_ids:
                                if leave.date == my_date:
                                    leave_type = leave.leave_type
                                    break

                            # Map leave type to style
                            if leave_type == 'holiday':
                                leave_style = holiday_style
                                status_label = 'Holiday'
                            elif leave_type == 'vacation':
                                leave_style = vacation_style
                                status_label = 'Vacation'
                            elif leave_type == 'medical':
                                leave_style = medical_leave_style
                                status_label = 'Medical Leave'
                            else:
                                status_label = 'Weekend' if week_of else 'Absent'

                            # Determine font to apply
                            if leave_type and not week_of:
                                font_to_set = leave_style
                            else:
                                font_to_set = ef
                            ####

                            # worksheet.merge_range(row, col, row, col + 1, attendance.employee_id.name, font_to_set_name)
                            worksheet.write(row, col + 0, sr_no_count, font_to_set)
                            sr_no_count += 1
                            #######
                            worksheet.write(row, col + 1, my_date.strftime("%A"), font_to_set)
                            worksheet.write(row, col + 2, my_date.strftime("%Y-%m-%d"), font_to_set)
                            worksheet.write(row, col + 3, status_label, font_to_set)
                            worksheet.write(row, col + 4, status_label, font_to_set)
                            worksheet.write(row, col + 5, ' ', font_to_set)
                            worksheet.write(row, col + 6, ' ', font_to_set)
                            worksheet.write(row, col + 7, ' ', font_to_set)

                            # Set working hours if not on leave or weekend
                            if leave_type in ['holiday', 'vacation', 'medical']:
                                worksheet.write(row, col + 8, '', font_to_set)
                                worksheet.write(row, col + 10, '', font_to_set)
                                worksheet.write(row, col + 11, '', font_to_set)
                                worksheet.write(row, col + 9, ' ', font_to_set)

                            else:
                                worksheet.write(row, col + 8,
                                                ' ',
                                                font_to_set)
                                worksheet.write(row, col + 10,
                                                ' ',
                                                font_to_set)
                                worksheet.write(row, col + 11,
                                                -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                                font_to_set)

                                worksheet.write(row, col + 9, -employee.resource_calendar_id.working_hours if not week_of else ' ', font_to_set)

                            #######
                            # worksheet.write(row, col + 1, my_date.strftime("%A"), font_to_set)
                            # worksheet.write(row, col + 2, my_date.strftime("%Y-%m-%d"), font_to_set)
                            # worksheet.write(row, col + 3, 'Absent' if not week_of else 'Weekend', font_to_set)
                            # worksheet.write(row, col + 4, 'Absent' if not week_of else 'Weekend', font_to_set)
                            # worksheet.write(row, col + 5, ' ', font_to_set)
                            # worksheet.write(row, col + 6, ' ', font_to_set)
                            # worksheet.write(row, col + 7, ' ', font_to_set)
                            # worksheet.write(row, col + 8, ' ', font_to_set)
                            # worksheet.write(row, col + 9,
                            #                 -employee.resource_calendar_id.working_hours if not week_of else ' ',
                            #                 font_to_set)
                            # worksheet.write(row, col + 10, ' ', font_to_set)
                            # worksheet.write(row, col + 11,
                            #                 -employee.resource_calendar_id.working_hours if not week_of else ' ',
                            #                 font_to_set)
                            if is_absent and not week_of:
                                total_absent_hours += employee.resource_calendar_id.working_hours or 8
                            row += 1
                            my_date += timedelta(days=1)
                        # ✅ Move to the next date to prevent infinite loop
                    p_on_weekend = True if my_date.weekday() == 4 else False
                    # worksheet.merge_range(row, col, row, col + 1, attendance.employee_id.name, present_name_weeend_style if p_on_weekend else font_left)
                    ###1

                    ef = font_center

                    # Initialize leave tracking
                    leave_type = None
                    leave_style = None
                    font_to_set = None

                    # Determine leave type on the given date
                    for leave in employee.leave_line_ids:
                        if leave.date == my_date:
                            leave_type = leave.leave_type
                            break
                    # Map leave type to style
                    if leave_type == 'holiday':
                        leave_style = holiday_present_style
                    elif leave_type == 'vacation':
                        leave_style = vacation_working_style
                    elif leave_type == 'medical':
                        leave_style = medical_leave_style

                    # Determine font to apply
                    # if leave_type and not p_on_weekend:
                    #     font_center = leave_style
                    # else:
                    #     font_center = ef
                    leave_ot = False
                    if leave_type == 'holiday' or leave_type == 'vacation':
                        leave_ot = True

                    # Merge vertically for the SR No. column if needed
                    if multi_punching_count > 1:
                        worksheet.merge_range(row, col + 0, row + multi_punching_count, col + 0, sr_no_count,present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    else:
                        worksheet.write(row, col + 0,sr_no_count,present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    sr_no_count += 1

                    if my_date.weekday() == 4:
                        total_friday_in_month += 1
                        total_day_in_months += 1
                    elif my_date.weekday() != 4:
                        total_day_in_months += 1
                    if attendance.check_in:
                        check_in = self.new_timezone(attendance.check_in)
                        if isinstance(check_in, str):  # Convert string to datetime
                            check_in = datetime.fromisoformat(check_in)
                        day_name = check_in.strftime("%A")
                        date_name = check_in.strftime("%Y-%m-%d")
                    elif attendance.check_out:
                        check_out = self.new_timezone(attendance.check_out)
                        if isinstance(check_out, str):  # Convert string to datetime
                            check_out = datetime.fromisoformat(check_out)
                        day_name = check_out.strftime("%A")
                        date_name = check_out.strftime("%Y-%m-%d")
                    else:
                        day_name = "N/A"
                        date_name = "N/A"
                    ###2
                    if multi_punching_count > 1:
                        worksheet.merge_range(row, col + 1, row + multi_punching_count, col + 1, day_name, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    else:
                        worksheet.write(row, col + 1, day_name, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    ###3
                    if multi_punching_count > 1:
                        worksheet.merge_range(row, col + 2, row + multi_punching_count, col + 2, date_name, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    else:
                        worksheet.write(row, col + 2, date_name, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    if attendance.check_in:
                        check_in = self.new_timezone(attendance.check_in)
                    else:
                        check_in = '***No Check In***'
                    worksheet.write(row, col + 3, check_in, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    if attendance.check_out:
                        check_out = self.new_timezone(attendance.check_out)
                    else:
                        check_out = '***No Check Out***'

                    worksheet.write(row, col + 4, check_out, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    diff_hours = int(attendance.worked_hours_ms)
                    diff_minutes = int((attendance.worked_hours_ms - diff_hours) * 60)
                    diff = f"{diff_hours:02}:{diff_minutes:02}"
                    worksheet.write(row, col + 5, diff, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    new_worked_hours = attendance.actual_worked_hours_ms
                    wrk_hours = int(new_worked_hours)
                    wrk_minutes = int((new_worked_hours - wrk_hours) * 60)
                    # wrk_hours = int(attendance.worked_hours)
                    # wrk_minutes = int((attendance.worked_hours - wrk_hours) * 60)
                    worked = f"{wrk_hours:02}:{wrk_minutes:02}"
                    if worked:
                        total_worked_hours += timedelta(hours=wrk_hours, minutes=wrk_minutes)
                    worksheet.write(row, col + 7, worked, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    ###9
                    if multi_punching_count > 1:
                        worksheet.merge_range(row, col + 8, row + multi_punching_count, col + 8, worked, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    else:
                        worksheet.write(row, col + 8, worked, present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))


                    break_time = datetime.strptime(diff, "%H:%M") - datetime.strptime(worked, "%H:%M")

                    # new_break_time = self.calculate_break_time(attendance)
                    def convert_decimal_to_time(decimal_hours):
                        hours = int(decimal_hours)
                        minutes = round((decimal_hours - hours) * 60)
                        return f"{hours:02d}:{minutes:02d}"  # Format as HH:MM

                    new_break_time = convert_decimal_to_time(attendance.break_time_ms) if attendance.break_time_ms else "00:00"
                    hours, remainder = divmod(total_worked_hours.seconds, 3600)
                    minutes = remainder // 60
                    formatted_break_time = f"{hours:02}:{minutes:02}"
                    worksheet.write(row, col + 6, new_break_time if new_break_time else formatted_break_time,
                                    present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    # Determine if the current day is Friday
                    check_in_date = datetime.strptime(check_in, "%Y-%m-%d %H:%M:%S") if attendance.check_in else None
                    is_friday = check_in_date and check_in_date.weekday() == 4  # 4 corresponds to Friday

                    # Write the shift hours based on whether it is Friday
                    ###10
                    if is_friday or leave_type in ['holiday', 'vacation', 'medical']:
                        if multi_punching_count > 1:
                            worksheet.merge_range(row, col + 9, row + multi_punching_count, col + 9, "00:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        else:
                            worksheet.write(row, col + 9, "00:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    elif employee.resource_calendar_id and employee.resource_calendar_id.working_hours:
                        if multi_punching_count > 1:
                            worksheet.merge_range(row, col + 9, row + multi_punching_count, col + 9, employee.resource_calendar_id.working_hours,
                                            present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        else:
                            worksheet.write(row, col + 9, employee.resource_calendar_id.working_hours,
                                            present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        total_working_hours += employee.resource_calendar_id.working_hours

                    else:
                        if multi_punching_count > 1:
                            worksheet.merge_range(row, col + 9, row + multi_punching_count, col + 9, "08:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        else:
                            worksheet.write(row, col + 9, "08:00", present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        total_working_hours += 8


                    def calculate_time(input_time, check_in, employee, overtime_hours_ms, shortfall_hours_ms,
                                       base_hours=8):
                        cumulative_overtime = timedelta(0)
                        cumulative_shortfall = timedelta(0)

                        # Convert input_time to timedelta
                        hours, minutes = map(int, input_time.split(":"))
                        input_timedelta = timedelta(hours=hours, minutes=minutes)
                        base_hours = 8
                        if employee and employee.resource_calendar_id and employee.resource_calendar_id.working_hours:
                            if int(employee.resource_calendar_id.working_hours) > 0:
                                base_hours = int(employee.resource_calendar_id.working_hours)

                        base_time = timedelta(hours=base_hours)

                        # Convert check_in to a datetime object to determine the day of the week
                        check_in_date = datetime.strptime(check_in, "%Y-%m-%d %H:%M:%S")
                        is_friday = check_in_date.weekday() == 4  # 4 corresponds to Friday

                        # Calculate adjusted time
                        adjusted_timedelta = input_timedelta - base_time

                        leave_type = None

                        # Determine leave type on the given date
                        for leave in employee.leave_line_ids:
                            if leave.date == my_date:
                                leave_type = leave.leave_type
                                break

                        leave_ot = False
                        if leave_type == 'holiday' or leave_type == 'vacation':
                            leave_ot = True

                        leave_type = None

                        for leave in employee.leave_line_ids:
                            if leave.date == my_date:
                                leave_type = leave.leave_type
                                break



                        if is_friday:
                            cumulative_overtime = input_time
                            classification = "Overtime"
                            return cumulative_overtime, classification
                        elif not is_friday and leave_ot:
                            cumulative_overtime = input_time
                            classification = "Overtime"
                            return cumulative_overtime, classification
                        else:
                            # Update cumulative counters and classify
                            if adjusted_timedelta > timedelta(0) or is_friday:
                                cumulative_overtime += adjusted_timedelta
                                classification = "Overtime"
                            else:
                                cumulative_shortfall -= adjusted_timedelta  # Subtract because adjusted_timedelta is negative
                                classification = "Shortfall"


                            # Extract hours and minutes
                            total_seconds = adjusted_timedelta.total_seconds()
                            total_minutes = total_seconds // 60
                            overtime_hours = total_minutes // 60
                            overtime_minutes = total_minutes % 60

                            # Apply corrected overtime rounding logic based on image
                            if overtime_hours == 0:
                                if 0 <= overtime_minutes <= 29:
                                    overtime_minutes = 0
                                elif 30 <= overtime_minutes <= 44:
                                    overtime_minutes = 30
                                elif 45 <= overtime_minutes <= 50:
                                    overtime_minutes = 45
                                elif 51 <= overtime_minutes <= 59:
                                    overtime_minutes = 0
                                    overtime_hours = 1
                            else:
                                if 0 <= overtime_minutes <= 29:
                                    overtime_minutes = 0
                                elif 30 <= overtime_minutes <= 44:
                                    overtime_minutes = 30
                                elif 45 <= overtime_minutes <= 50:
                                    overtime_minutes = 45
                                elif 51 <= overtime_minutes <= 59:
                                    overtime_minutes = 0
                                    overtime_hours += 1

                            formatted_adjusted_time = f"{int(overtime_hours):02}:{int(overtime_minutes):02}"

                            return formatted_adjusted_time, classification

                    # Function to format cumulative time for display
                    def format_cumulative_time(timedelta_obj):
                        hours, remainder = divmod(timedelta_obj.seconds, 3600)
                        minutes = remainder // 60
                        return f"{hours:02}:{minutes:02}"

                    # Example usage:
                    adjusted_time, classification = calculate_time(worked, check_in, employee,attendance.overtime_hours_ms ,attendance.shortfall_hours_ms)  # Friday

                    def round_time(adjusted_time):
                        hours, minutes = map(int, adjusted_time.split(":"))
                        total_minutes = hours * 60 + minutes
                        new_hours = total_minutes // 60
                        remaining_minutes = total_minutes % 60
                        if new_hours <= 0:
                            if 0 <= remaining_minutes <= 29:
                                new_minutes = 0
                            elif 30 <= remaining_minutes <= 44:
                                new_minutes = 30
                            elif 45 <= remaining_minutes <= 50:
                                new_minutes = 45
                            else:  # 51 to 59
                                new_hours += 1
                                new_minutes = 0
                        elif new_hours > 0:
                            if 0 <= remaining_minutes <= 14:
                                new_minutes = 0
                            elif 15 <= remaining_minutes <= 45:
                                new_minutes = 30
                            elif 46 <= remaining_minutes <= 59:
                                new_minutes = 0
                                new_hours += 1  # Round up to the next hour

                        return f"{new_hours:02}:{new_minutes:02}"
                    ###11
                    if adjusted_time and my_date.weekday() == 4:

                        rounded_time = round_time(adjusted_time)
                        # time_spited = rounded_time.split(':')
                        # total_overtime_hours_rounded += float(time_spited[0]) or 0
                        # total_overtime_minutes_rounded += float(time_spited[1]) or 0
                        if classification == "Overtime":
                            total_ot_list.append(rounded_time)

                        if multi_punching_count > 1:
                            worksheet.merge_range(row, col + 10, row + multi_punching_count, col + 10, rounded_time if classification == "Overtime" else "00:00",
                                            present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        else:
                            worksheet.write(row, col + 10, rounded_time if classification == "Overtime" else "00:00",
                                            present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    else:
                        rounded_time = round_time(adjusted_time)
                        # time_spited = rounded_time.split(':')
                        # total_overtime_hours_rounded += float(time_spited[0]) or 0
                        # total_overtime_minutes_rounded += float(time_spited[1]) or 0
                        if classification == "Overtime":
                            total_ot_list.append(rounded_time)
                        if multi_punching_count > 1:
                            worksheet.merge_range(row, col + 10, row + multi_punching_count, col + 10,
                                                  rounded_time if classification == "Overtime" else "00:00",
                                                  present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                        else:
                            worksheet.write(row, col + 10, rounded_time if classification == "Overtime" else "00:00",
                                            present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    hours1, minutes1 = map(int, total_overtime.split(":"))
                    hours2, minutes2 = map(int,
                                           adjusted_time.split(":") if classification == "Overtime" else "00:00".split(":"))

                    time1_delta = timedelta(hours=hours1, minutes=minutes1)
                    time2_delta = timedelta(hours=hours2, minutes=minutes2)
                    # Add the two timedelta objects
                    total_time = time1_delta + time2_delta

                    # Convert back to "HH:MM" format
                    total_hours, remainder = divmod(total_time.seconds, 3600)
                    total_minutes = remainder // 60
                    total_overtime = f"{total_time.days * 24 + total_hours:02}:{total_minutes:02}"

                    ###12
                    if multi_punching_count > 1:
                        worksheet.merge_range(row, col + 11, row + multi_punching_count, col + 11,
                                              '-' + adjusted_time if classification == "Shortfall" else "00:00",
                                              present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))
                    else:
                        worksheet.write(row, col + 11, '-' + adjusted_time if classification == "Shortfall" else "00:00",
                                        present_weeend_style if p_on_weekend else (leave_style if leave_style else font_center))

                    if classification == "Shortfall" and adjusted_time:
                        hours, minutes = map(int, adjusted_time.split(":"))
                        total_shortfall_hours += timedelta(hours=hours, minutes=minutes)




                    my_date += timedelta(days=1)
                    row += 1
                    ########################
                    def format_datetime(dt, user_tz):
                        if not dt:
                            return ''
                        if not isinstance(dt, datetime):
                            return str(dt)
                        local_tz = timezone(user_tz or 'UTC')
                        dt = UTC.localize(dt).astimezone(local_tz)
                        return dt.strftime('%Y-%m-%d %H:%M:%S')


                    def format_hours(hours):
                        return '{:02}:{:02}'.format(int(hours), int((hours % 1) * 60)) if hours is not None else ''

                    user_tz = self.env.context.get('tz') or 'UTC'

                    if attendance.multiple_checkin_ids and len(attendance.multiple_checkin_ids) > 1:
                        for punching in attendance.multiple_checkin_ids:
                            # worksheet.write(row, col + 0, ' ')  # SR No
                            # worksheet.write(row, col + 1, ' ')  # Day
                            # worksheet.write(row, col + 2, ' ')  # Date
                            check_in = self.new_timezone(punching.check_in)
                            check_out = self.new_timezone(punching.check_out)
                            worksheet.write(row, col + 3, check_in or ' ',
                                            multiple_shift_font)
                            worksheet.write(row, col + 4, check_out or ' ',
                                            multiple_shift_font)
                            worksheet.write(row, col + 5, format_hours(punching.worked_hours) or ' ', multiple_shift_font)
                            worksheet.write(row, col + 6, format_hours(punching.break_time) or ' ', multiple_shift_font)
                            worksheet.write(row, col + 7, format_hours(punching.actual_worked_hours) or ' ',
                                            multiple_shift_font)
                            # worksheet.write(row, col + 8, ' ')
                            # worksheet.write(row, col + 9, ' ')
                            # worksheet.write(row, col + 10, ' ')
                            # worksheet.write(row, col + 11, ' ')
                            row += 1

                if my_date < last_date:
                    while my_date < last_date:
                        font_to_set = working_day_absent_style
                        font_to_set_name = name_working_day_absent_style
                        week_of = my_date and my_date.weekday() == 4
                        if week_of:
                            font_to_set = absent_weeend_style
                            font_to_set_name = name_weekend_absent_style
                            total_friday_in_month += 1
                            total_day_in_months += 1
                        else:
                            total_day_in_months += 1
                            total_absent_in_month += 1

                        # worksheet.merge_range(row, col, row, col + 1, attendance.employee_id.name, font_to_set_name)
                        worksheet.write(row, col + 0, sr_no_count, font_to_set)
                        sr_no_count += 1

                        ###
                        ef = font_to_set
                        leave_type = None
                        leave_style = None
                        font_to_set = None

                        # Determine leave type on the given date
                        for leave in employee.leave_line_ids:
                            if leave.date == my_date:
                                leave_type = leave.leave_type
                                break

                        # Apply style if present and check-in exists
                        if leave_type and attendance and attendance.check_in:
                            if leave_type == 'holiday':
                                leave_style = holiday_style
                            elif leave_type == 'vacation':
                                leave_style = vacation_style
                            elif leave_type == 'medical':
                                leave_style = medical_leave_style

                        # Apply font if it's a leave day and not weekend
                        if leave_type and not week_of:
                            font_to_set = leave_style
                        else:
                            font_to_set = ef

                        # Set label depending on leave or weekend
                        if week_of:
                            status_label = 'Weekend'
                        elif leave_type == 'holiday':
                            status_label = 'Holiday'
                        elif leave_type == 'vacation':
                            status_label = 'Vacation'
                        elif leave_type == 'medical':
                            status_label = 'Medical Leave'
                        else:
                            status_label = 'Absent'

                        ###
                        worksheet.write(row, col + 1, my_date.strftime("%A"), font_to_set)
                        worksheet.write(row, col + 2, my_date.strftime("%Y-%m-%d"), font_to_set)
                        worksheet.write(row, col + 3, status_label, font_to_set)
                        worksheet.write(row, col + 4, status_label, font_to_set)
                        worksheet.write(row, col + 5, ' ', font_to_set)
                        worksheet.write(row, col + 6, ' ', font_to_set)
                        worksheet.write(row, col + 7, ' ', font_to_set)

                        # If leave_type in holiday/vacation/medical, set '' in hours, else working hours
                        if leave_type in ['holiday', 'vacation', 'medical']:
                            worksheet.write(row, col + 8, '', font_to_set)
                            worksheet.write(row, col + 10, '', font_to_set)
                            worksheet.write(row, col + 11, '', font_to_set)

                        else:
                            worksheet.write(row, col + 8,
                                            -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                            font_to_set)
                            worksheet.write(row, col + 10,
                                            -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                            font_to_set)
                            worksheet.write(row, col + 11,
                                            -employee.resource_calendar_id.working_hours if not week_of else ' ',
                                            font_to_set)

                        worksheet.write(row, col + 9, ' ', font_to_set)

                        ###
                        # worksheet.write(row, col + 1, my_date.strftime("%A"), font_to_set)  # Day name
                        # worksheet.write(row, col + 2, my_date.strftime("%Y-%m-%d"), font_to_set)
                        # worksheet.write(row, col + 3, 'Absent' if not week_of else 'Weekend', font_to_set)
                        # worksheet.write(row, col + 4, 'Absent' if not week_of else 'Weekend', font_to_set)
                        # worksheet.write(row, col + 5, ' ', font_to_set)
                        # worksheet.write(row, col + 6, ' ', font_to_set)
                        # worksheet.write(row, col + 7, ' ', font_to_set)
                        # worksheet.write(row, col + 8, ' ', font_to_set)
                        # worksheet.write(row, col + 9, -employee.resource_calendar_id.working_hours if not week_of else ' ',
                        #                 font_to_set)
                        # worksheet.write(row, col + 10, ' ', font_to_set)
                        # worksheet.write(row, col + 11, -employee.resource_calendar_id.working_hours if not week_of else ' ',
                        #                 font_to_set)
                        if not week_of:
                            total_absent_hours += employee.resource_calendar_id.working_hours or 8
                        my_date += timedelta(days=1)
                        row += 1
                ####
                new_over_time = ''
                # Convert '91:57' to total minutes
                total_overtime_minutes = sum(
                    (int(ots.split(':')[0]) * 60 + int(ots.split(':')[1]))
                    if ots and ':' in ots and '-' not in ots.split(':')[0] else 0
                    for ots in total_ot_list
                )

                # hours, minutes = map(int, "91:57".split(':'))
                # total_overtime_minutes = (total_overtime_hours_rounded * 60) + total_overtime_minutes_rounded
                # Convert 112 hours to minutes
                total_absent_minutes = total_absent_hours * 60  # 112*60 = 6720 minutes

                # Subtract minutes
                # remaining_minutes = total_overtime_minutes - total_absent_minutes
                remaining_minutes = total_overtime_minutes

                # Convert back to hours and minutes
                remaining_hours = remaining_minutes // 60
                remaining_mins = abs(remaining_minutes % 60)  # Use abs() to avoid negative minutes
                new_over_time = f"{remaining_hours}:{remaining_mins:02d}"
                # Format result
                # if remaining_minutes < 0:
                #     new_over_time = f"-{abs(remaining_hours)}:{remaining_mins:02d}"  # Handle negative time
                # else:
                #     new_over_time = f"{remaining_hours}:{remaining_mins:02d}"

                worksheet.write(row, 10, f"Total: {new_over_time}", font_center)
                # worksheet.write(row, 10, f"Total: {total_overtime}", font_center)
                worksheet.write(row, 9, f"Total: {total_working_hours}", font_center)
                # Convert timedelta to total hours and minutes
                total_hours = total_shortfall_hours.seconds // 3600  # Extract hours
                total_minutes = (total_shortfall_hours.seconds % 3600) // 60  # Extract minutes
                #####
                total_seconds = total_worked_hours.total_seconds()
                total_hours = int(total_seconds // 3600)
                total_minutes = int((total_seconds % 3600) // 60)
                formatted_worked_time = f"{total_hours:02}:{total_minutes:02}"
                # worksheet.write(row, 11, f"Total: {total_hours:02}:{total_minutes:02}", font_center)
                worksheet.write(row, 8, f"Total: {formatted_worked_time}", font_center)

                # calculation for net overtime
                net_total_overtime_mins = total_overtime_minutes - total_absent_minutes
                nt_total_seconds = net_total_overtime_mins * 60
                nt_total_hours = int(nt_total_seconds // 3600)
                nt_total_minutes = int((nt_total_seconds % 3600) // 60)
                nt_formatted_ot_time = f"{nt_total_hours:02}:{nt_total_minutes:02}"

                # totals
                worksheet.merge_range(row + 2, 0, row + 2, 1, 'Total Days in month', font_bold_left)
                worksheet.merge_range(row + 3, 0, row + 3, 1, 'Total Friday', font_bold_left)
                worksheet.merge_range(row + 4, 0, row + 4, 1, 'Working days in a month', font_bold_left)
                worksheet.merge_range(row + 5, 0, row + 5, 1, 'Absent / Shortfall', font_bold_left)
                worksheet.merge_range(row + 6, 0, row + 6, 1, 'Total Absent Hours', font_bold_left)
                worksheet.merge_range(row + 7, 0, row + 7, 1, 'Total Overtime on month', font_bold_left)
                worksheet.merge_range(row + 8, 0, row + 8, 1, 'Net Overtime', font_bold_left)

                # worksheet.merge_range(row, col, row, col + 1, attendance.employee_id.name, font_to_set_name)

                worksheet.write(row + 2, 2, total_day_in_months, font_bold_center)
                worksheet.write(row + 3, 2, total_friday_in_month, font_bold_center)
                worksheet.write(row + 4, 2, total_day_in_months - total_friday_in_month, font_bold_center)
                worksheet.write(row + 5, 2, total_absent_in_month, font_bold_center)
                worksheet.write(row + 6, 2, total_absent_hours, font_bold_center)
                worksheet.write(row + 7, 2, new_over_time if new_over_time else '', font_bold_center)
                worksheet.write(row + 8, 2, nt_formatted_ot_time or '', font_bold_center)

        workbook.close()
        xlsx_data = output.getvalue()

        return [fl, xlsx_data]

    def new_timezone(self, time):
        """
        Convert a naive UTC datetime to the user's local timezone and return as string.

        Args:
            time (datetime): A naive datetime object in UTC.

        Returns:
            str: The datetime converted to the user's timezone in '%Y-%m-%d %H:%M:%S' format.
        """
        try:
            user_timezone = self.env.user.tz or str(pytz.UTC)
            local_tz = pytz.timezone(user_timezone)
            converted_datetime = (
                pytz.UTC.localize(time, is_dst=False)
                .astimezone(local_tz)
            )
            return datetime.strftime(converted_datetime, "%Y-%m-%d %H:%M:%S")
        except Exception as e:
            raise UserError(f"Unable to convert timezone: {str(e)}")

    def tnis_usr_timezone(self, dt_obj):
        """
        Convert a naive datetime (assumed UTC) to a naive datetime in user's local timezone.

        Args:
            dt_obj (datetime): Naive datetime object assumed to be in UTC.

        Returns:
            datetime: Naive datetime adjusted to user's timezone.
        """
        try:
            tz_name = self.env.user.tz
            user_tz = pytz.timezone(tz_name) if tz_name else pytz.UTC
            localized_dt = (
                pytz.UTC.localize(dt_obj.replace(tzinfo=None), is_dst=False)
                .astimezone(user_tz)
                .replace(tzinfo=None)
            )
            return localized_dt
        except Exception as e:
            raise UserError(f"Failed to adjust datetime to user timezone: {str(e)}")

    def tni_user_utc_tz(self, dt_obj):
        """
        Convert a naive datetime in user's local timezone to a naive UTC datetime.

        Args:
            dt_obj (datetime): Naive datetime in user's timezone.

        Returns:
            datetime: Naive datetime in UTC.
        """
        try:
            tz_name = self.env.user.tz
            user_tz = pytz.timezone(tz_name) if tz_name else pytz.UTC
            utc_dt = (
                user_tz.localize(dt_obj.replace(tzinfo=None), is_dst=False)
                .astimezone(pytz.UTC)
                .replace(tzinfo=None)
            )
            return utc_dt
        except Exception as e:
            raise UserError(f"Failed to convert datetime to UTC: {str(e)}")

    def dps_from_to_timezone(self, dt_obj):
        """
        Convert a naive datetime (assumed UTC) to user's timezone and keep tzinfo removed.

        Args:
            dt_obj (datetime): Naive datetime object in UTC.

        Returns:
            datetime: Naive datetime adjusted to user's timezone.
        """
        try:
            tz_name = self.env.user.tz
            user_tz = pytz.timezone(tz_name) if tz_name else pytz.UTC
            converted_dt = (
                pytz.UTC.localize(dt_obj.replace(tzinfo=None), is_dst=False)
                .astimezone(user_tz)
                .replace(tzinfo=None)
            )
            return converted_dt
        except Exception as e:
            raise UserError(f"Failed to apply user timezone: {str(e)}")

    def user_tz_convert(self, time_str):
        """
        Convert a datetime string from user's local timezone to UTC format compatible with Odoo.

        Args:
            time_str (str): Datetime string in '%Y-%m-%d %H:%M:%S' format.

        Returns:
            str: Datetime in UTC as Odoo-compatible string.
        """
        try:
            naive_dt = datetime.strptime(str(time_str), '%Y-%m-%d %H:%M:%S')

            parsed_dt = datetime.strptime(naive_dt.strftime('%Y-%m-%d %H:%M:%S'), '%Y-%m-%d %H:%M:%S')

            user_tz = pytz.timezone(self.env.user.tz or 'GMT')
            local_dt = user_tz.localize(parsed_dt, is_dst=False)

            utc_dt = local_dt.astimezone(pytz.UTC)
            formatted_utc = utc_dt.strftime("%Y-%m-%d %H:%M:%S")

            final_dt = datetime.strptime(formatted_utc, "%Y-%m-%d %H:%M:%S")
            return fields.Datetime.to_string(final_dt)

        except ValueError:
            raise UserError("Invalid datetime format. Please provide a valid date in 'YYYY-MM-DD HH:MM:SS' format.")
        except Exception as e:
            raise UserError(f"Timezone conversion failed: {str(e)}")
    ################################################################################################################



    def export_employee_attendance_from_logs(self, logs):
        """
        Generate an Excel report for attendance logs within the specified date range.

        Args:
            logs (list): List of attendance log records containing:
                         - employee_id.name: Employee Name
                         - user_punch_time: Datetime of the punch
                         - status: Attendance status (0=Check In, 1=Check Out, else=Punched)
                         - device: Device name or identifier

        Returns:
            list: [filename, binary_excel_data]
                  filename (str): The generated Excel file name
                  binary_excel_data (bytes): The Excel file content as binary
        """
        try:
            start_date = datetime.strptime(str(self.report_date_start_from), '%Y-%m-%d %H:%M:%S').date()
            end_date = datetime.strptime(str(self.report_date_end_to), '%Y-%m-%d %H:%M:%S').date()

            start_day, start_month, start_year = start_date.strftime('%d'), start_date.strftime(
                '%B'), start_date.strftime('%Y')
            end_day, end_month, end_year = end_date.strftime('%d'), end_date.strftime('%B'), end_date.strftime('%Y')

            report_title = f"Employee Attendance Log From {start_day}-{start_month}-{start_year} To {end_day}-{end_month}-{end_year}"
            file_name = f"{report_title} ({datetime.today()}).xlsx"

            output_stream = io.BytesIO()
            workbook = xlsxwriter.Workbook(output_stream)

            worksheet = workbook.add_worksheet('Attendance Logs')
            worksheet.set_landscape()  # Landscape mode for better width

            header_format = workbook.add_format({
                'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter',
                'bg_color': '#4472C4', 'font_color': 'white', 'border': 1
            })
            subheader_format = workbook.add_format({
                'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter',
                'bg_color': '#D9E1F2', 'border': 1
            })
            text_left_format = workbook.add_format({
                'align': 'left', 'border': 1, 'font_size': 11, 'bg_color': '#F8F9FA'
            })
            text_center_format = workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 11
            })
            text_center_highlight_format = workbook.add_format({
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 11, 'bg_color': '#E2EFDA'
            })

            worksheet.set_column('A:E', 25)
            worksheet.set_column('F:XFD', None, None, {'hidden': True})

            worksheet.merge_range('A1:E1', report_title, header_format)
            worksheet.set_row(0, 28)

            row, col = 2, 0
            worksheet.merge_range(row, col, row + 1, col + 1, "Employee Name", subheader_format)
            worksheet.merge_range(row, col + 2, row + 1, col + 2, "Punching Time", subheader_format)
            worksheet.merge_range(row, col + 3, row + 1, col + 3, "Status", subheader_format)
            worksheet.merge_range(row, col + 4, row + 1, col + 4, "Device", subheader_format)

            row += 2
            for log in logs:
                worksheet.merge_range(row, col, row, col + 1, log.employee_id.name or 'Unknown', text_left_format)

                user_punch_time = self.new_timezone(log.user_punch_time) if log.user_punch_time else '***No Status***'
                worksheet.write(row, col + 2, user_punch_time, text_center_format)

                if log.status == "0":
                    status_text = 'Check In'
                elif log.status == "1":
                    status_text = 'Check Out'
                else:
                    status_text = 'Punched'
                worksheet.write(row, col + 3, status_text, text_center_highlight_format)

                worksheet.write(row, col + 4, log.device or 'N/A', text_center_format)

                row += 1

            workbook.close()
            xlsx_data = output_stream.getvalue()

            return [file_name, xlsx_data]

        except Exception as e:
            # Raise professional error message
            raise UserError(f"An error occurred while generating the attendance report: {str(e)}")


    def calculate_break_time(self, attendance):
        if not attendance or not attendance.check_in or not attendance.check_out:
            return "00:00"

        check_in_utc = attendance.check_in
        check_out_utc = attendance.check_out
        employee_id = attendance.employee_id
        working_schedule = employee_id.resource_calendar_id

        if not working_schedule:
            return "00:00"

        # Get user's timezone
        user_tz = self.env.user.tz or 'UTC'
        local_tz = timezone(user_tz)

        check_in = check_in_utc.astimezone(local_tz)
        check_out = check_out_utc.astimezone(local_tz)

        check_in_day = check_in.weekday()  # Monday = 0, Sunday = 6
        check_out_day = check_out.weekday()

        relevant_attendances = [
            att for att in working_schedule.attendance_ids
            if att.day_period == 'lunch' and att.dayofweek in {str(check_in_day), str(check_out_day)}
        ]

        total_break_time = 0.0
        for att in relevant_attendances:
            start_time = att.hour_from
            end_time = att.hour_to

            break_start = check_in.replace(hour=int(start_time), minute=int((start_time % 1) * 60))
            break_end = check_in.replace(hour=int(end_time), minute=int((end_time % 1) * 60))

            if check_in <= break_start and break_end <= check_out:
                break_duration = (break_end - break_start).total_seconds() / 60  # Convert to minutes
                total_break_time += break_duration

        hours, minutes = divmod(int(total_break_time), 60)

        return f"{hours:02}:{minutes:02}"

