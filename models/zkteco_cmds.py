# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import models, fields

class DeviceCommand(models.Model):
    _name = 'zkteco.dcmmand'
    _description = 'Device Command'

    name = fields.Char(string='Command Name', required=True)
    device_id = fields.Many2one('zkteco.device.setting', string='Device', required=True)
    employee_id = fields.Many2one('hr.employee', "Employee")
    status = fields.Selection([
        ('pending', 'Pending'),
        ('executed', 'Executed'),
        ('success', 'Success'),
        ('failed', 'Failed'),
    ], string='Status', default='pending')
    pin = fields.Integer('PIN')
    execution_log = fields.Text(string='Execution Log')

