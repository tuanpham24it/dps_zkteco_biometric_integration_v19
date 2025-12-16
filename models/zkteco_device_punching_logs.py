# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import api, fields, models, _
from datetime import datetime
from odoo.exceptions import UserError, ValidationError
from convertdate import islamic


# ======================================================
# ZKTeco Device Logs
# ======================================================

class ZktecoDeviceLogs(models.Model):
    _name = 'zkteco.device.logs'
    _description = 'ZKTeco Device Logs'
    _order = 'user_punch_time desc'
    _rec_name = 'user_punch_time'

    status = fields.Selection(
        [
            ('0', 'Check In'),
            ('1', 'Check Out'),
            ('2', 'Punched')
        ],
        string='Status'
    )

    user_punch_time = fields.Datetime(string='Punching Time')
    user_punch_calculated = fields.Boolean(string='Punch Calculated', default=False)
    device = fields.Char(string='Device')

    company_id = fields.Many2one(
        'res.company',
        string='Company',
        default=lambda self: self.env.company,
        readonly=True
    )

    zketco_duser_id = fields.Many2one(
        'zkteco.attendance.machine',
        string="Device User"
    )

    employee_id = fields.Many2one(
        'hr.employee',
        related='zketco_duser_id.employee_id',
        store=True
    )

    employee_code = fields.Char(string='Employee Code')
    employee_department = fields.Char(
        related='zketco_duser_id.employee_id.department_id.name',
        string='Department'
    )
    employee_name = fields.Char(
        related='zketco_duser_id.employee_id.name',
        string='Employee Name'
    )

    weekday_name = fields.Char(
        string="Weekday",
        compute="_compute_weekday_name",
        store=True
    )

    status_number = fields.Char(string="Status Number")
    number = fields.Char(string="Number")
    timestamp = fields.Integer(string="Timestamp")
    punch_status_in_string = fields.Char(string="Status String")

    # --------------------------------------------------
    # CREATE (ODOO 19 SAFE)
    # --------------------------------------------------

    @api.model_create_multi
    def create(self, vals_list):
        records = super().create(vals_list)

        hr_attendance = self.env['hr.attendance']

        for record in records:
            employee = record.employee_id
            punch_time = record.user_punch_time

            if not employee or not punch_time:
                continue

            last_attendance = hr_attendance.search(
                [('employee_id', '=', employee.id)],
                order='check_in desc',
                limit=1
            )

            if not last_attendance or last_attendance.check_out:
                hr_attendance.create({
                    'employee_id': employee.id,
                    'check_in': punch_time,
                })
                record.status = '0'  # Check In
            else:
                if punch_time > last_attendance.check_in:
                    last_attendance.write({'check_out': punch_time})
                    record.status = '1'  # Check Out
                else:
                    record.status = '2'  # Old punch

        return records

    # --------------------------------------------------
    # UNLINK PROTECTION
    # --------------------------------------------------

    def unlink(self):
        processed = self.filtered(lambda r: r.user_punch_calculated)
        if processed:
            raise UserError(_("You cannot delete processed attendance logs."))
        return super().unlink()

    # --------------------------------------------------
    # COMPUTE WEEKDAY
    # --------------------------------------------------

    @api.depends('user_punch_time')
    def _compute_weekday_name(self):
        weekdays = [
            "Thứ Hai", "Thứ Ba", "Thứ Tư",
            "Thứ Năm", "Thứ Sáu", "Thứ Bảy", "Chủ Nhật"
        ]
        for rec in self:
            rec.weekday_name = (
                weekdays[rec.user_punch_time.weekday()]
                if rec.user_punch_time else False
            )


# ======================================================
# HR ATTENDANCE EXTENSION
# ======================================================

class HrAttendance(models.Model):
    _inherit = "hr.attendance"

    punch_date = fields.Date(string='Punch Date')

    is_multiple_shift = fields.Boolean(
        string="Is Multiple Shift",
        compute='_compute_multiple_shifts',
        store=True
    )

    break_time_ms = fields.Float(compute="_compute_ms_fields", store=True)
    worked_hours_ms = fields.Float(compute="_compute_ms_fields", store=True)
    actual_worked_hours_ms = fields.Float(compute="_compute_ms_fields", store=True)
    overtime_hours_ms = fields.Float(compute="_compute_ms_fields", store=True)
    shortfall_hours_ms = fields.Float(compute="_compute_ms_fields", store=True)

    shortfall = fields.Float(string='Shortfall Hours', compute='_compute_shortfall', store=True)

    leave_type = fields.Selection([
        ('none', 'None'),
        ('holiday', 'Holiday'),
        ('medical', 'Medical Leave'),
        ('vacation', 'Vacation'),
    ], default='none', required=True)

    # --------------------------------------------------
    # MULTIPLE SHIFT CONFIG
    # --------------------------------------------------

    @api.depends()
    def _compute_multiple_shifts(self):
        param = self.env['ir.config_parameter'].sudo().get_param(
            'dps_zkteco_biometric_integration.multiple_shift'
        )
        enabled = param in ['True', 'true', '1']
        for rec in self:
            rec.is_multiple_shift = enabled

    def _get_multiple_shift_status(self):
        param = self.env['ir.config_parameter'].sudo().get_param(
            'dps_zkteco_biometric_integration.multiple_shift'
        )
        return param in ['True', 'true', '1']

    # --------------------------------------------------
    # CREATE (ODOO 19 SAFE)
    # --------------------------------------------------

    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if 'is_multiple_shift' not in vals:
                vals['is_multiple_shift'] = self._get_multiple_shift_status()
        return super().create(vals_list)

    def write(self, values):
        if 'is_multiple_shift' not in values:
            values['is_multiple_shift'] = self._get_multiple_shift_status()
        return super().write(values)

    # --------------------------------------------------
    # RAMADAN CALENDAR
    # --------------------------------------------------

    def is_in_ramadan(self, date):
        hijri = islamic.from_gregorian(date.year, date.month, date.day)
        return hijri[1] == 9

    def _get_employee_calendar(self):
        self.ensure_one()
        if self.employee_id and self.check_in:
            if self.is_in_ramadan(self.check_in.date()):
                return self.employee_id.ramadan_resource_calendar_id
        return super()._get_employee_calendar()

    # --------------------------------------------------
    # COMPUTE SHORTFALL
    # --------------------------------------------------

    @api.depends('worked_hours', 'employee_id')
    def _compute_shortfall(self):
        for rec in self:
            rec.shortfall = 0.0
            if rec.employee_id and rec.worked_hours:
                calendar = rec._get_employee_calendar()
                if calendar:
                    working_hours = sum(calendar.attendance_ids.filtered(
                        lambda a: a.dayofweek == str(rec.check_in.weekday())
                    ).mapped('duration_hours'))
                    if working_hours > rec.worked_hours:
                        rec.shortfall = 2 * (working_hours - rec.worked_hours)

    # --------------------------------------------------
    # MULTI SHIFT AGGREGATE
    # --------------------------------------------------

    @api.depends(
        'multiple_checkin_ids.break_time',
        'multiple_checkin_ids.worked_hours',
        'multiple_checkin_ids.actual_worked_hours'
    )
    def _compute_ms_fields(self):
        for rec in self:
            rec.break_time_ms = sum(rec.multiple_checkin_ids.mapped('break_time'))
            rec.worked_hours_ms = sum(rec.multiple_checkin_ids.mapped('worked_hours'))
            rec.actual_worked_hours_ms = sum(rec.multiple_checkin_ids.mapped('actual_worked_hours'))

            working_hours = rec.employee_id.resource_calendar_id.working_hours or 0
            diff = rec.actual_worked_hours_ms - working_hours

            rec.overtime_hours_ms = max(diff, 0)
            rec.shortfall_hours_ms = abs(min(diff, 0))

    # --------------------------------------------------
    # CHECK IN / OUT DIFF
    # --------------------------------------------------

    check_in_check_out_difference = fields.Float(
        string='Punching Difference',
        compute='check_in_check_out_diff'
    )

    def check_in_check_out_diff(self):
        for rec in self:
            if rec.check_in and rec.check_out:
                delta = rec.check_out - rec.check_in
                if delta.total_seconds() < 0:
                    raise ValidationError(_("Check-out cannot be earlier than check-in."))
                rec.check_in_check_out_difference = delta.total_seconds() / 3600
            else:
                rec.check_in_check_out_difference = 0.0

    # --------------------------------------------------
    # UNLINK RESET LOG STATE
    # --------------------------------------------------

    def unlink(self):
        logs = self.env['zkteco.device.logs']
        for rec in self:
            related_logs = logs.search([
                ('employee_id', '=', rec.employee_id.id),
                '|',
                ('user_punch_time', '=', rec.check_in),
                ('user_punch_time', '=', rec.check_out),
            ])
            related_logs.write({'user_punch_calculated': False})
        return super().unlink()
