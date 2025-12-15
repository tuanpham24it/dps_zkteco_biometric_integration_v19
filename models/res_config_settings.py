# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import fields, models


class ResConfigSettingsInherit(models.TransientModel):
    """
    Inherits system configuration settings to add ZKTeco-specific options.

    Provides:
    - Minimal Attendance: Enables a mode where attendance records are stored in minimal form.
    - Multiple Shift: Allows multiple shift handling for employees.
    """
    _inherit = 'res.config.settings'

    multiple_shift = fields.Boolean(
        string='Multiple Shift',
        store=True,
        copy=False,
        config_parameter='dps_zkteco_biometric_integration.multiple_shift'
    )

    minimal_attendance = fields.Boolean(
        string='User Minimal Attendance',
        config_parameter='dps_zkteco_biometric_integration.minimal_attendance'
    )
