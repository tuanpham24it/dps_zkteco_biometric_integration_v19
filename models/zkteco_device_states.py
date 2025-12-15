# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import models, fields, api
from odoo.exceptions import ValidationError, UserError


class ZktecoDeviceState(models.Model):
    """
    Model representing attendance states received from biometric devices.

    This model is used to categorize and define the state of employee
    check-ins and check-outs that are captured by integrated biometric devices.
    Each record corresponds to a specific device and activity type.
    """

    _name = 'zkteco.device.states'
    _description = "Biometric Device Attendance State"


    name = fields.Char(
        string='State Name',
        required=True,
        help="The descriptive name of the attendance state. Example: 'Morning Check-In'."
    )

    code = fields.Char(
        string='Code',
        required=True,
        help="Unique technical code to identify the state internally. Example: 'CHK_IN_AM'."
    )

    device_id = fields.Many2one('zkteco.device.setting', string="Device", required=True)

    description = fields.Text(
        string='Description',
        help="Additional details or notes describing the purpose or usage of this state."
    )

    activity_type = fields.Selection(
        [
            ('check_in', 'Check-In'),
            ('check_out', 'Check-Out')
        ],
        string="Activity Type",
        help="Defines whether this state represents an employee Check-In or Check-Out event."
    )

