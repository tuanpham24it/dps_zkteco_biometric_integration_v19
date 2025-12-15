# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import api, fields, models, _
from odoo.exceptions import ValidationError

class ZktecoDeviceFingerprints(models.Model):
    """
    Model to store fingerprint templates for employees on ZKTeco biometric devices.

    Each record represents a fingerprint template for a specific employee linked
    to a particular device. The template is required for employee authentication
    on biometric devices.
    """
    _name = 'zkteco.device.fingerprints'
    _description = 'Employee Fingerprint Templates for ZKTeco Devices'

    employee_id = fields.Many2one(
        'hr.employee',
        string='Employee',
        related='zketco_duser_id.employee_id',
        store=True,
        readonly=True,
        help="Employee associated with this fingerprint template."
    )

    name = fields.Char(
        related='employee_id.name',
        string="Employee Name",
        store=True,
        copy=False,
        readonly=True,
        help="Name of the employee linked to this fingerprint template."
    )

    device_id = fields.Many2one('zkteco.device.setting', string="Device")

    zketco_duser_id = fields.Many2one(
        'zkteco.attendance.machine',
        string='Device User',
        help="The device user entry associated with this employee on the biometric device."
    )

    template_data = fields.Binary(
        string='Template Data',
        required=True,
        help="Binary data representing the employee's fingerprint template."
    )
