# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import models, fields
from datetime import datetime


class DeviceStampLog(models.Model):
    """
    Model to store synchronization stamp logs from biometric devices.

    Each record represents a synchronization attempt with a ZKTeco (or similar)
    biometric attendance device, including:
    - The device reference
    - The timestamp of the sync
    - The raw log text returned
    - The last processed stamp (used for incremental syncs)
    """
    _name = 'device.stamp.logs'

    name = fields.Char(string='Stamp Name')

    log_date = fields.Datetime(
        string='Log Date',
        default=lambda self: datetime.now()
    )

    log_text = fields.Text("Log Text")

    device_id = fields.Many2one(
        'zkteco.device.setting',
        string='Biometric Attendance Device',
        required=True
    )

    stamp = fields.Integer("Stamp")


class DeviceOperationStampLogs(models.Model):


    _name = 'device.operation.stamplogs'

    name = fields.Char(
        string='Log Title',
        help='A short descriptive title for the log entry.'
    )
    log_date = fields.Datetime(
        string='Log Timestamp',
        default=lambda self: datetime.now(),
        help='The date and time when the log entry was created.'
    )
    log_text = fields.Text(
        string='Log Details',
        help='Detailed description or notes regarding this log entry.'
    )
    device_id = fields.Many2one(
        'zkteco.device.setting',
        string='Biometric Device',
        required=True,
        help='The biometric attendance device from which this log originates.'
    )
    opStamp = fields.Integer(
        string='Operation Stamp',
        help='Numerical stamp value representing the operation performed.'
    )

