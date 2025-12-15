from odoo import models, fields, api
from odoo.exceptions import ValidationError
from datetime import timedelta, datetime, time
import pytz


class EmployeeLeaveWizard(models.TransientModel):
    _name = 'employee.leave.wizard'
    _description = 'Employee Leave Wizard'

    leave_type = fields.Selection([
        ('none', 'None'),
        ('holiday', 'Holiday'),
        ('medical', 'Medical Leave'),
        ('vacation', 'Vacation'),
    ], string="Leave Type", default='none', required=True)

    start_date = fields.Date(string="Start Date", required=True)
    end_date = fields.Date(string="End Date", required=True)
    employee_ids = fields.Many2many('hr.employee', string="Employees", required=True)
    description = fields.Char(string="Description")

    create_date = fields.Datetime(string="Created On", readonly=True, default=fields.Datetime.now)
    user_id = fields.Many2one('res.users', string="Created By", default=lambda self: self.env.user, readonly=True)
    paid_medical_leave = fields.Boolean(string='Paid')
    att_start_date = fields.Datetime(string="Start", compute="_compute_att_dates", inverse="_inverse_att_dates",
                                     store=True)
    att_end_date = fields.Datetime(string="End", compute="_compute_att_dates", inverse="_inverse_att_dates", store=True)

    @api.depends('paid_medical_leave')
    def _compute_att_dates(self):
        user_tz_name = self.env.user.tz or 'UTC'
        user_tz = pytz.timezone(user_tz_name)

        for rec in self:
            if rec.paid_medical_leave:
                now_utc = fields.Datetime.now()
                now_user = now_utc.astimezone(pytz.utc).astimezone(user_tz)
                today = now_user.date()

                start_local = user_tz.localize(datetime.combine(today, time(8, 0)))
                end_local = user_tz.localize(datetime.combine(today, time(17, 0)))

                rec.att_start_date = start_local.astimezone(pytz.utc).replace(tzinfo=None)
                rec.att_end_date = end_local.astimezone(pytz.utc).replace(tzinfo=None)
            else:
                rec.att_start_date = False
                rec.att_end_date = False

    def _inverse_att_dates(self):
        user_tz_name = self.env.user.tz or 'UTC'
        user_tz = pytz.timezone(user_tz_name)

        for rec in self:
            if rec.paid_medical_leave:
                # If paid is True but dates are not set, assign default 8 AM – 5 PM
                if not rec.att_start_date or not rec.att_end_date:
                    now_utc = fields.Datetime.now()
                    now_user = now_utc.astimezone(pytz.utc).astimezone(user_tz)
                    today = now_user.date()

                    start_local = user_tz.localize(datetime.combine(today, time(8, 0)))
                    end_local = user_tz.localize(datetime.combine(today, time(17, 0)))

                    rec.att_start_date = start_local.astimezone(pytz.utc).replace(tzinfo=None)
                    rec.att_end_date = end_local.astimezone(pytz.utc).replace(tzinfo=None)
            else:
                # If user unset paid_medical_leave, clear dates too
                rec.att_start_date = False
                rec.att_end_date = False



    def action_create_leave_lines(self):
        if self.end_date < self.start_date:
            raise ValidationError("End date must be greater than or equal to start date.")

        leave_dates = []
        current_date = self.start_date
        while current_date <= self.end_date:
            leave_dates.append(current_date)
            current_date += timedelta(days=1)

        for employee in self.employee_ids:
            for date in leave_dates:
                already_exists = self.env['employee.leave.line'].search_count([
                    ('employee_id', '=', employee.id),
                    ('date', '=', date),
                ])

                if not already_exists:
                    self.env['employee.leave.line'].create({
                        'employee_id': employee.id,
                        'date': date,
                        'leave_type': self.leave_type,
                        'description': self.description,
                        'att_start_date': self.att_start_date,
                        'att_end_date': self.att_end_date,
                        'paid_medical_leave': self.paid_medical_leave,
                    })

        return {'type': 'ir.actions.client', 'tag': 'reload'}

    def action_cancel(self):
        return {'type': 'ir.actions.act_window_close'}
