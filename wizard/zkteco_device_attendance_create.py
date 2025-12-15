# -*- coding: utf-8 -*-

from datetime import datetime, timedelta
from odoo import api, fields, models, _
from pytz import timezone, UTC
# from odoo.addons.resource.models.utils import

# Customized by Tunn
from odoo.fields import Domain
from odoo.tools.intervals import Intervals

# # Domain
# try:
#     from odoo.fields import Domain        # Odoo 19
# except ImportError:
#     Domain = None                         # Odoo <=18 không có Domain

# # Intervals
# try:
#     from odoo.tools.intervals import Intervals   # Odoo 19
# except ImportError:
#     from odoo.addons.resource.models.utils import Intervals  # Odoo <=18

from datetime import datetime, time


class ZktecoCalculationWizard(models.TransientModel):
    _name = 'zkteco.calculation.wizard'

    def device_user_check_in_out(self, emp_id, time):
        """
        Retrieve the active attendance record for an employee.

        This method checks whether the specified employee currently has an
        open attendance entry (i.e., a record with `check_out` still empty).

        Args:
            emp_id (int): The ID of the employee to check.
            time (datetime): The current time or the time of check (currently unused in logic).

        Returns:
            int or None: The ID of the open attendance record if found, otherwise None.
        """
        attendances = self.env['hr.attendance'].search([
            ('employee_id', '=', emp_id),
            ('check_out', '=', False)
        ])

        if attendances:
            return attendances.id


    def adjust_check_in_out_times(self, check_in, check_out):
        """
        Adjust employee check-in and check-out times based on company rules and user timezone.

        This method performs the following operations:
            1. Converts `check_in` and `check_out` from UTC to the user's local timezone.
            2. Adjusts the `check_in` time:
                - If it's earlier than 8:15 AM, or before 1:00 AM, it is reset to 8:00 AM.
            3. Adjusts the `check_out` time (if provided):
                - Only applies adjustment if check-out is after 5:00 PM.
                - Rounds the minutes according to the following rules:
                    * 00-29 → 00
                    * 30-44 → 30
                    * 45-50 → 45
                    * 51-59 → Next hour (with minutes set to 00)
            4. Converts both times back to UTC (without timezone info).

        Args:
            check_in (datetime): Original check-in time in UTC.
            check_out (datetime or None): Original check-out time in UTC (if available).

        Returns:
            tuple: (adjusted_check_in, adjusted_check_out) as naive datetime objects in UTC.
        """

        user_tz = timezone(self.env.user.tz) if self.env.user.tz else UTC

        check_in_local = check_in.replace(tzinfo=UTC).astimezone(user_tz)
        check_out_local = check_out.replace(tzinfo=UTC).astimezone(user_tz) if check_out else None

        if check_in_local.time() < time(8, 15) or check_in_local.time() < time(1, 0):
            check_in_local = check_in_local.replace(hour=8, minute=0, second=0, microsecond=0)

        if check_out_local and check_out_local.time() > time(17, 0):
            hour = check_out_local.hour
            minute = check_out_local.minute

            if 0 <= minute <= 29:
                minute = 0
            elif 30 <= minute <= 44:
                minute = 30
            elif 45 <= minute <= 50:
                minute = 45
            elif 51 <= minute <= 59:
                hour += 1
                minute = 0

            check_out_local = check_out_local.replace(hour=hour, minute=minute, second=0, microsecond=0)

        check_in = check_in_local.astimezone(UTC).replace(tzinfo=None)
        check_out = check_out_local.astimezone(UTC).replace(tzinfo=None) if check_out_local else None

        return check_in, check_out

    def calculate_attendance(self):
        """
        Calculate and adjust attendance records based on attendance logs and company settings.

        This method processes attendance logs from biometric devices or other sources and creates or updates
        corresponding HR attendance records in Odoo. It handles:
            - Timezone conversions
            - Adjusting check-in and check-out times
            - Applying minimal attendance rules (if configured)
            - Assigning leave types for specific dates
            - Handling both single-shift and multi-shift configurations
            - Pairing multiple punches into check-in/check-out pairs

        Logic is divided into two paths:
            1. Single shift scenario (multiple_shift = False)
            2. Multi-shift scenario (multiple_shift = True)

        Helper Functions:
            adjust_check_in_out_times: Normalizes check-in and check-out times based on business rules.
            get_leave_type_for_date: Determines the leave type for an employee on a given date.

        Raises:
            No direct exceptions raised here, but writes and creates may raise Odoo ORM exceptions if
            constraints fail.

        Returns:
            None
        """

        def adjust_check_in_out_times(check_in, check_out):
            """
            Adjusts check-in and check-out times according to predefined rules:
            - Convert UTC times to user's timezone
            - Normalize check-in before 8:15 AM to exactly 8:00 AM
            - Round check-out times after 5 PM to the nearest 0, 30, 45, or next hour

            Args:
                check_in (datetime): Original check-in time in UTC
                check_out (datetime or None): Original check-out time in UTC

            Returns:
                tuple: (adjusted_check_in, adjusted_check_out) in UTC without tzinfo
            """
            user_tz = timezone(self.env.user.tz) if self.env.user.tz else UTC

            check_in_local = check_in.replace(tzinfo=UTC).astimezone(user_tz)
            check_out_local = check_out.replace(tzinfo=UTC).astimezone(user_tz) if check_out else None

            if check_in_local.time() < time(8, 15) or check_in_local.time() < time(1, 0):
                check_in_local = check_in_local.replace(hour=8, minute=0, second=0, microsecond=0)

            if check_out_local and check_out_local.time() > time(17, 0):
                minute = check_out_local.minute
                hour = check_out_local.hour

                if 0 <= minute <= 29:
                    minute = 0
                elif 30 <= minute <= 44:
                    minute = 30
                elif 45 <= minute <= 50:
                    minute = 45
                elif 51 <= minute <= 59:
                    if hour < 23:
                        hour += 1
                        minute = 0
                    else:
                        minute = 59

                check_out_local = check_out_local.replace(hour=hour, minute=minute, second=0, microsecond=0)

            check_in = check_in_local.astimezone(UTC).replace(tzinfo=None)
            check_out = check_out_local.astimezone(UTC).replace(tzinfo=None) if check_out_local else None

            return check_in, check_out

        def get_leave_type_for_date(employee, dt):
            """
            Get the leave type for the given employee on the specified date.

            Args:
                employee (record): hr.employee record
                dt (datetime): Punch time (UTC)

            Returns:
                str: Leave type string (or 'none' if no leave found)
            """
            user_tz = timezone(self.env.user.tz) if self.env.user.tz else UTC
            local_date = dt.replace(tzinfo=UTC).astimezone(user_tz).date()
            leave_line = self.env['employee.leave.line'].search([
                ('employee_id', '=', employee.id),
                ('date', '=', local_date)
            ], limit=1)
            return leave_line.leave_type if leave_line else 'none'

        param_val = self.env['ir.config_parameter'].sudo().get_param('dps_zkteco_biometric_integration.multiple_shift')

        if param_val in [False, 'False', 'false', '0', 0, None, '']:
            minimal_attendance = self.env['ir.config_parameter'].sudo().get_param(
                'dps_zkteco_biometric_integration.minimal_attendance')
            hr_attendance = self.env['hr.attendance']
            today = datetime.now(UTC)

            domain = [('user_punch_time', '<=', today), ('user_punch_calculated', '=', False)]
            attendance_log = self.env['zkteco.device.logs'].search(domain).sorted('user_punch_time')
            employee_list = []

            for log in attendance_log.filtered(lambda x: x.employee_id):
                user_tz = timezone(self.env.user.tz) if self.env.user.tz else UTC
                punch_time_utc = log.user_punch_time.replace(tzinfo=UTC)
                punch_time_local = punch_time_utc.astimezone(user_tz)

                morning_start = punch_time_local.replace(hour=7, minute=45, second=0, microsecond=0)
                morning_end = punch_time_local.replace(hour=8, minute=15, second=0, microsecond=0)
                if morning_start <= punch_time_local <= morning_end:
                    punch_time_local = punch_time_local.replace(hour=8, minute=0, second=0)

                punch_time = punch_time_local.astimezone(UTC).replace(tzinfo=None)

                if minimal_attendance:
                    attendance = self.env['hr.attendance'].search([
                        ('employee_id', '=', log.employee_id.id),
                        ('punch_date', '=', log.user_punch_time.date())
                    ])
                    if attendance:
                        if punch_time > attendance.check_in:
                            _, adjusted_check_out = adjust_check_in_out_times(attendance.check_in, punch_time)
                            try:
                                attendance.write({'check_out': adjusted_check_out})
                            except:
                                continue
                    else:
                        last_attendance_before_check_out = self.env['hr.attendance'].search([
                            ('employee_id', '=', log.employee_id.id),
                            ('check_out', '=', False)
                        ], order='check_in desc', limit=1)
                        if last_attendance_before_check_out:
                            check_out_time = last_attendance_before_check_out.check_in.replace(hour=23, minute=59,
                                                                                               second=59)
                            if check_out_time > last_attendance_before_check_out.check_in:
                                _, adjusted_check_out = adjust_check_in_out_times(
                                    last_attendance_before_check_out.check_in, check_out_time)
                                try:
                                    last_attendance_before_check_out.write({'check_out': adjusted_check_out})
                                except:
                                    continue

                        adjusted_check_in, _ = adjust_check_in_out_times(punch_time, None)
                        hr_attendance.create({
                            'employee_id': log.employee_id.id,
                            'check_in': adjusted_check_in,
                            'punch_date': punch_time.date(),
                            'leave_type': get_leave_type_for_date(log.employee_id, punch_time),
                        })

                else:
                    if log.employee_id.id in employee_list:
                        attd = self.device_user_check_in_out(log.employee_id.id, punch_time)
                        attendance = self.env['hr.attendance'].browse(attd)
                        if attendance and punch_time > attendance.check_in:
                            _, adjusted_check_out = adjust_check_in_out_times(attendance.check_in, punch_time)
                            try:
                                attendance.write({'check_out': adjusted_check_out})
                            except:
                                continue
                        employee_list.remove(log.employee_id.id)
                    else:
                        attd = self.device_user_check_in_out(log.employee_id.id, punch_time)
                        attendance = self.env['hr.attendance'].browse(attd)
                        if attendance and attendance.check_out is False:
                            if punch_time > attendance.check_in:
                                _, adjusted_check_out = adjust_check_in_out_times(attendance.check_in, punch_time)
                                try:
                                    attendance.write({'check_out': adjusted_check_out})
                                except:
                                    continue
                        else:
                            adjusted_check_in, _ = adjust_check_in_out_times(punch_time, None)
                            hr_attendance.create({
                                'employee_id': log.employee_id.id,
                                'check_in': adjusted_check_in,
                                'leave_type': get_leave_type_for_date(log.employee_id, punch_time),
                            })
                            employee_list.append(log.employee_id.id)

                log.user_punch_calculated = True

        else:
            minimal_attendance = self.env['ir.config_parameter'].sudo().get_param(
                'dps_zkteco_biometric_integration.minimal_attendance')
            hr_attendance = self.env['hr.attendance']
            user_tz = timezone(self.env.user.tz) if self.env.user.tz else UTC

            today = datetime.now(UTC)
            domain = [('user_punch_time', '<=', today), ('user_punch_calculated', '=', False)]
            attendance_log = self.env['zkteco.device.logs'].search(domain).sorted('user_punch_time')

            logs_by_employee = {}

            for log in attendance_log.filtered(lambda x: x.employee_id):
                punch_time_utc = log.user_punch_time.replace(tzinfo=UTC)
                punch_time_local = punch_time_utc.astimezone(user_tz)
                punch_time = punch_time_local.astimezone(UTC).replace(tzinfo=None)

                key = (log.employee_id.id, punch_time.date())
                logs_by_employee.setdefault(key, []).append((log, punch_time))

            for (employee_id, punch_date), log_entries in logs_by_employee.items():
                # Fetch existing attendance for the day
                attendance = self.env['hr.attendance'].search([
                    ('employee_id', '=', employee_id),
                    ('punch_date', '=', punch_date)
                ], limit=1)

                if attendance:
                    all_logs = self.env['zkteco.device.logs'].search([
                        ('employee_id', '=', employee_id),
                        ('user_punch_time', '>=', datetime.combine(punch_date, datetime.min.time())),
                        ('user_punch_time', '<=', datetime.combine(punch_date, datetime.max.time())),
                    ]).sorted('user_punch_time')

                    log_entries = []
                    for log in all_logs:
                        punch_time_utc = log.user_punch_time.replace(tzinfo=UTC)
                        punch_time_local = punch_time_utc.astimezone(user_tz)
                        punch_time = punch_time_local.astimezone(UTC).replace(tzinfo=None)
                        log_entries.append((log, punch_time))

                log_entries.sort(key=lambda x: x[1])
                punch_times = [pt for (_, pt) in log_entries]

                if len(punch_times) < 2:
                    continue

                check_in = punch_times[0]
                check_out = punch_times[-1]
                adjusted_check_in, adjusted_check_out = adjust_check_in_out_times(check_in, check_out)

                leave_type = get_leave_type_for_date(self.env['hr.employee'].browse(employee_id), check_in)
                vals = {
                    'check_in': adjusted_check_in,
                    'check_out': adjusted_check_out if adjusted_check_in != adjusted_check_out else False,
                    'leave_type': leave_type,
                    'punch_date': punch_date,
                }

                if attendance:
                    attendance.write(vals)
                else:
                    vals['employee_id'] = employee_id
                    attendance = self.env['hr.attendance'].create(vals)

                attendance.multiple_checkin_ids.unlink()
                count = 1
                i = 0
                while i < len(punch_times):
                    punch_in = punch_times[i]
                    punch_out = punch_times[i + 1] if i + 1 < len(punch_times) else None
                    self.env['multiple.punch'].create({
                        'attendance_id': [(6, 0, [attendance.id])],
                        'count': count,
                        'check_in': punch_in,
                        'check_out': punch_out,
                    })
                    count += 1
                    i += 2

                for (log, _) in log_entries:
                    log.user_punch_calculated = True



class MultiplePuching(models.Model):
    _name = 'multiple.punch'
    _description = 'Multiple Punch'

    attendance_id = fields.Many2many('hr.attendance', string='Attendance', store=True, copy=False)
    employee_id = fields.Many2one('hr.employee', string='Employee', related='attendance_id.employee_id', copy=False,
                                  store=True)
    count = fields.Integer(string='Count', store=True, copy=False)
    check_in = fields.Datetime(string="Check In", tracking=True, store=True, copy=False)
    check_out = fields.Datetime(string="Check Out", tracking=True, store=True, copy=False)
    break_time = fields.Float(string="Break Time", compute="_compute_work_hours_and_breaks", store=True)
    worked_hours = fields.Float(string="Total Worked Hours", compute="_compute_work_hours_and_breaks", store=True,
                                copy=False)
    actual_worked_hours = fields.Float(string="Actual Worked Hours", compute="_compute_work_hours_and_breaks",
                                       store=True, copy=False)

    @api.depends("check_in", "check_out", "employee_id")
    def _compute_work_hours_and_breaks(self):
        for attendance in self:
            worked_hours = 0.0
            break_time = 0.0
            actual_worked_hours = 0.0

            if attendance.check_in and attendance.check_out:
                delta = attendance.check_out - attendance.check_in
                worked_hours = delta.total_seconds() / 3600.0  # in hours

            if not attendance.check_in or not attendance.check_out or not attendance.employee_id.resource_calendar_id:
                attendance.worked_hours = worked_hours
                attendance.break_time = 0.0
                attendance.actual_worked_hours = worked_hours
                continue

            user_tz = self.env.user.tz or "UTC"
            local_tz = timezone(user_tz)
            check_in = attendance.check_in.astimezone(local_tz)
            check_out = attendance.check_out.astimezone(local_tz)

            check_in_day = str(check_in.weekday())
            check_out_day = str(check_out.weekday())

            working_schedule = attendance.employee_id.resource_calendar_id
            relevant_attendances = [
                att for att in working_schedule.attendance_ids
                if att.day_period == "lunch" and att.dayofweek in {check_in_day, check_out_day}
            ]

            for att in relevant_attendances:
                start_time = att.hour_from
                end_time = att.hour_to

                break_start = check_in.replace(hour=int(start_time), minute=int((start_time % 1) * 60))
                break_end = check_in.replace(hour=int(end_time), minute=int((end_time % 1) * 60))

                if check_in <= break_start <= check_out and check_in <= break_end <= check_out:
                    break_duration = (break_end - break_start).total_seconds() / 3600.0
                    break_time += break_duration

            actual_worked_hours = worked_hours - break_time

            # Assign all computed values
            attendance.worked_hours = round(worked_hours, 2)
            attendance.break_time = round(break_time, 2)
            attendance.actual_worked_hours = round(actual_worked_hours, 2)


class HrAttendance(models.Model):
    _inherit = "hr.attendance"

    break_time = fields.Float(compute="_compute_calculated_attendance_break_time", store=True)
    multiple_checkin_ids = fields.One2many('multiple.punch', 'attendance_id', string='Multiple Checkins')

    @api.depends("check_in", "check_out", "employee_id")
    def _compute_calculated_attendance_break_time(self):
        """
        Compute the total break time (in hours) for an employee within a given attendance record.

        This calculation is based on:
        - The employee's working schedule (`resource_calendar_id`).
        - The attendance record's check-in and check-out times.
        - Lunch break rules defined in the employee's schedule.

        Steps:
        1. Validate that check-in, check-out, and working schedule exist.
        2. Convert check-in and check-out times to the user's timezone.
        3. Identify lunch periods defined in the schedule for the relevant weekdays.
        4. Calculate total break duration that falls within the attendance period.
        5. Store the result as `break_time` in hours, rounded to two decimal places.

        Assumptions:
        - Breaks considered are those marked as `day_period='lunch'` in attendance rules.
        - Only breaks fully within the attendance window are counted.

        Raises:
        - No explicit errors, but invalid or missing data will result in 0.0 break time.
        """
        for attendance in self:
            if not attendance.check_in or not attendance.check_out or not attendance.employee_id.resource_calendar_id:
                attendance.break_time = 0.0
                continue

            check_in_utc = attendance.check_in
            check_out_utc = attendance.check_out
            employee = attendance.employee_id
            working_schedule = employee.resource_calendar_id

            user_tz = self.env.user.tz or "UTC"
            local_tz = timezone(user_tz)

            check_in = check_in_utc.astimezone(local_tz)
            check_out = check_out_utc.astimezone(local_tz)

            check_in_day = str(check_in.weekday())
            check_out_day = str(check_out.weekday())

            relevant_attendances = [
                att for att in working_schedule.attendance_ids
                if att.day_period == "lunch" and att.dayofweek in {check_in_day, check_out_day}
            ]

            total_break_time = 0.0  # Initialize total break time accumulator

            for att in relevant_attendances:
                start_time = att.hour_from
                end_time = att.hour_to

                break_start = check_in.replace(hour=int(start_time), minute=int((start_time % 1) * 60))
                break_end = check_in.replace(hour=int(end_time), minute=int((end_time % 1) * 60))

                if check_in <= break_start <= check_out and check_in <= break_end <= check_out:
                    break_duration = (break_end - break_start).total_seconds() / 3600  # Convert to hours
                    total_break_time += break_duration

            attendance.break_time = round(total_break_time, 2)

    @api.depends('check_in', 'check_out')
    def _compute_worked_hours(self):
        """
        Compute the total worked hours for an attendance record.

        This calculation:
        - Considers the employee's calendar and its time zone.
        - Excludes lunch break intervals from the total worked time.
        - Converts check-in and check-out times to the employee's calendar timezone.

        Steps:
        1. Validate the presence of check-in, check-out, and employee record.
        2. Determine the employee's working calendar and corresponding timezone.
        3. Calculate intervals for attendance, excluding lunch breaks.
        4. Convert the total duration from seconds to hours and assign to `worked_hours`.

        Behavior:
        - If mandatory data (check-in, check-out, employee) is missing, `worked_hours` will be set to `False`.

        Notes:
        - Relies on `_get_employee_calendar()` and `_employee_attendance_intervals()` for calendar and break handling.
        - Does not raise errors; missing or invalid calendar defaults to `False` (handled gracefully).

        Raises:
        - No explicit exceptions. If data is missing, worked hours remain unset (False).
        """
        for attendance in self:
            if attendance.check_out and attendance.check_in and attendance.employee_id:
                calendar = attendance._get_employee_calendar() or attendance.employee_id.resource_calendar_id
                if not calendar:
                    attendance.worked_hours = 0.0
                    continue

                try:
                    tz = timezone(calendar.tz)
                except Exception:
                    attendance.worked_hours = 0.0
                    continue

                check_in_tz = attendance.check_in.astimezone(tz)
                check_out_tz = attendance.check_out.astimezone(tz)

                lunch_intervals = attendance.employee_id._employee_attendance_intervals(
                    check_in_tz, check_out_tz, lunch=True
                )

                attendance_intervals = Intervals([(check_in_tz, check_out_tz, attendance)]) - lunch_intervals

                total_seconds = sum((interval[1] - interval[0]).total_seconds() for interval in attendance_intervals)
                attendance.worked_hours = total_seconds / 3600.0
            else:
                attendance.worked_hours = False
