# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import api, fields, models, _
from odoo.exceptions import UserError

from odoo import models, fields, _
from odoo.exceptions import UserError


class HrEmployee(models.Model):
    _inherit = 'hr.employee'

    biometric_device_ids = fields.One2many(
        'zkteco.attendance.machine',
        'employee_id',
        string='Biometric Devices',
        help='List of biometric devices linked to this employee.'
    )

    # Customized by Tunn
    # arabic_name = fields.Char()
    # ramadan_resource_calendar_id = fields.Many2one(
    #     'resource.calendar', string='Ramadan Working Hours', check_company=True)
    arabic_name = fields.Char(
        string="Arabic Name",
        groups="base.group_user" # Internal user
    )

    ramadan_resource_calendar_id = fields.Many2one(
        'resource.calendar',
        string='Ramadan Working Hours',
        check_company=True,
        groups="base.group_user" # Internal user
    )


    leave_line_ids = fields.One2many('employee.leave.line', 'employee_id', string="Leave Lines", tracking=True)



    def create_export_command(self, device_id):

        existing_command = self.env['zkteco.dcmmand'].sudo().search([
            ('employee_id', '=', self.id),
            ('device_id', '=', device_id.id),
            ('name', '=', 'DATA'),
            ('status', '=', 'pending')
        ])
        if existing_command:
            raise UserError(_("A pending 'DATA' command already exists for this employee on this device."))

        all_zketco_duser_ids = self.env['zkteco.attendance.machine'].sudo().search([]).mapped('zkteco_device_attend_id')
        int_user_ids_list = [int(uid) for uid in all_zketco_duser_ids]

        pending_pin_list = self.env['zkteco.dcmmand'].search([('status', '!=', 'success')]).mapped('pin')

        combined_pin_list = int_user_ids_list + pending_pin_list
        next_pin = max(combined_pin_list) + 1 if combined_pin_list else 1

        command = self.env['zkteco.dcmmand'].sudo().create({
            'name': 'DATA',
            'device_id': device_id.id,
            'employee_id': self.id,
            'status': 'pending',
            'pin': next_pin,
        })

        card_number = self.barcode if self.barcode else "0000000000"
        command.execution_log = (
            f"C:{command.id}:DATA USER PIN={next_pin} "
            f"Name={self.name} Pri=0 Passwd= Card=[{card_number}] Grp=1 TZ=0000000000000000\n"
        )

    def employee_del_command(self, device_id):

        existing_command = self.env['zkteco.dcmmand'].sudo().search([
            ('employee_id', '=', self.id),
            ('device_id', '=', device_id.id),
            ('name', '=', 'DEL'),
            ('status', '=', 'pending')
        ])
        if existing_command:
            raise UserError(_("A pending delete command already exists for this employee on this device."))

        matched_device = self.biometric_device_ids.filtered(lambda d: d.device_id == device_id)
        if not matched_device:
            raise UserError(_("The employee is not registered on the selected device."))

        command = self.env['zkteco.dcmmand'].sudo().create({
            'name': 'DEL',
            'device_id': device_id.id,
            'employee_id': self.id,
            'status': 'pending',
            'pin': matched_device.zkteco_device_attend_id,
        })

        command.execution_log = f"C:{command.id}:DATA DEL_USER PIN={matched_device.zkteco_device_attend_id} \n"

    def update_export_command(self, device_id):

        existing_command = self.env['zkteco.dcmmand'].sudo().search([
            ('employee_id', '=', self.id),
            ('device_id', '=', device_id.id),
            ('name', '=', 'UPDATE'),
            ('status', '=', 'pending')
        ])
        if existing_command:
            raise UserError(_("A pending 'UPDATE' command already exists for this employee on this device."))

        matched_device = self.biometric_device_ids.filtered(lambda d: d.device_id == device_id)
        if not matched_device:
            raise UserError(_("The employee is not registered on the selected device."))

        command = self.env['zkteco.dcmmand'].sudo().create({
            'name': 'UPDATE',
            'device_id': device_id.id,
            'employee_id': self.id,
            'status': 'pending',
            'pin': matched_device.zkteco_device_attend_id,
        })

        command.execution_log = f"C:{command.id}:DATA USER PIN={matched_device.zkteco_device_attend_id} Name={self.name} \n"


class ZktecoAttendanceMachine(models.Model):
    """
    Model representing individual biometric attendance device users.

    Stores the link between an employee and their user ID on a specific
    biometric attendance device.
    """

    _name = 'zkteco.attendance.machine'
    _rec_name = 'zkteco_device_attend_id'

    zkteco_device_username = fields.Char(
        string='Device Username',
        help='The username associated with the employee on the biometric device.'
    )
    employee_id = fields.Many2one(
        'hr.employee',
        string='Employee',
        help='The employee linked to this biometric device user.'
    )
    zkteco_device_attend_id = fields.Char(
        string='Device User ID',
        required=True,
        help='Unique user ID on the ZKTeco attendance device.'
    )
    device_id = fields.Many2one(
        'zkteco.device.setting',
        string='ZKTeco Device',
        required=True,
        help='The specific biometric attendance device for this user.'
    )
    employee_badge_barcode_id = fields.Char(
        string='Badge ID',
        store=True,
        related='employee_id.barcode',
        help='Employee badge barcode, automatically fetched from the employee record.'
    )


class ResourceCalendarInherit(models.Model):
    """
    Inherited resource calendar model to add custom working hours field.

    This allows tracking of working hours specifically configured for each calendar.
    """
    _inherit = 'resource.calendar'

    working_hours = fields.Float(
        string='Working Hours',
        copy=False,
        store=True,
        help='Total working hours defined for this resource calendar.'
    )
