# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import api, fields, models
from odoo.exceptions import UserError, ValidationError


class ZKTecoDeviceEventLog(models.Model):
    """
    Model: ZKTecoDeviceLog
    ----------------------
    This model represents system and user operation logs generated
    by biometric devices (e.g., ZKTeco). These logs help track
    important events such as power cycles, enrollment activities,
    deletions, and system resets.

    Useful for auditing and monitoring device activity.
    """
    _name = 'zkteco.device.event.log'   # Better name for clarity
    _description = "ZKTeco Device Operation Logs"

    device_id = fields.Many2one(
        'zkteco.device.setting',
        string="Device",
        help="Reference to the biometric device that generated this log."
    )
    log_code = fields.Char(
        string='Log Code',
        help="Unique identifier for the log event as provided by the device."
    )
    description = fields.Selection(
        [
            ('-1', 'N/A'),
            ('0', 'Power On'),
            ('1', 'Power Off'),
            ('2', 'Authentication Failure'),
            ('3', 'Alarm'),
            ('4', 'Enter Menu'),
            ('5', 'Change Settings'),
            ('6', 'Enroll Fingerprint'),
            ('7', 'Enroll Password'),
            ('8', 'Enroll HID Card'),
            ('9', 'Delete User'),
            ('10', 'Delete Fingerprint'),
            ('11', 'Delete Password'),
            ('12', 'Delete RF Card'),
            ('13', 'Clear Data'),
            ('14', 'Create MF Card'),
            ('15', 'Enroll MF Card'),
            ('16', 'Register MF Card'),
            ('17', 'Delete MF Card Registration'),
            ('18', 'Clear MF Card Content'),
            ('19', 'Move Enrollment Data to Card'),
            ('20', 'Copy Data from Card to Machine'),
            ('21', 'Set Time'),
            ('22', 'Factory Reset'),
            ('23', 'Delete Entry/Exit Records'),
            ('24', 'Clear Administrator Permissions'),
            ('25', 'Modify Access Group Settings'),
            ('26', 'Modify User Access Settings'),
            ('27', 'Modify Access Time Zones'),
            ('28', 'Modify Unlocking Combination Settings'),
            ('29', 'Unlock'),
            ('30', 'Enroll New User'),
            ('31', 'Change Fingerprint Properties'),
            ('32', 'Forced Alarm'),
            ('33', 'Doorbell Call'),
            ('34', 'Anti-submarine'),
            ('35', 'Delete Attendance Photo'),
            ('36', 'Modify User Other Information'),
            ('37', 'Holiday'),
            ('38', 'Restore Data')
        ],
        string='Description Log Code',
        help="Type of log event captured from the device."
    )
    operator = fields.Char(
        string='Operator',
        help="The user or system responsible for triggering the log event."
    )
    op_time = fields.Datetime(
        string='Created On',
        help="Timestamp when the log was created on the device."
    )
    value_1 = fields.Char(
        string='Value 1',
        help="Optional parameter provided by the device for additional context."
    )
    value_2 = fields.Char(
        string='Value 2',
        help="Optional parameter provided by the device for additional context."
    )
    value_3 = fields.Char(
        string='Value 3',
        help="Optional parameter provided by the device for additional context."
    )
    reserved = fields.Char(
        string='Reserved',
        help="Reserved field for future device-specific data."
    )
    opStamp = fields.Integer(
        string='OpStamp',
        help="Operation stamp number used for synchronization and ordering."
    )

    @api.constrains('log_code', 'device_id')
    def _check_unique_log_per_device(self):
        """
        Ensures that the same log code cannot be duplicated
        for the same device. This avoids redundant log entries.
        """
        for log in self:
            duplicate = self.search([
                ('log_code', '=', log.log_code),
                ('device_id', '=', log.device_id.id),
                ('id', '!=', log.id)
            ], limit=1)
            if duplicate:
                raise ValidationError(
                    f"The log with code '{log.log_code}' already exists for the selected device."
                )
