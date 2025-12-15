# -*- coding: utf-8 -*-
from odoo import api, fields, models, _
from ..zk import ZK
from odoo.exceptions import ValidationError, UserError


class ZktecoDeviceWizard(models.TransientModel):
    """
    Wizard for performing actions on ZKTeco biometric devices.

    Provides options to:
    - Update device data
    - Scan logs
    - Remove employees or data
    """
    _name = 'zkteco.device.wizard'
    _description = 'ZKTeco Device Wizard'

    zkteco_device_id = fields.Many2one(
        'zkteco.device.setting',
        string='ZKTeco Device'
    )

    zkteco_device_ids = fields.Many2many(
        'zkteco.device.setting',
        string='ZKTeco Devices',
        required=True
    )

    operation_type = fields.Selection(
        [
            ('update', 'Update'),
            ('scan', 'Scan'),
            ('remove', 'Remove')
        ],
        string="Operation Type"
    )

    def validate_biometric_update(self):

        employee_id = self._context.get('active_id')
        employee = self.env['hr.employee'].browse(employee_id)

        for biometric_device in self.zkteco_device_ids:

            if biometric_device.is_adms:
                zketco_duser_id = self.env['zkteco.attendance.machine'].search([
                    ('employee_id', '=', employee.id),
                    ('device_id', '=', biometric_device.id),
                ])

                if not zketco_duser_id:
                    employee.create_export_command(biometric_device)
                else:
                    employee.update_export_command(biometric_device)

            else:
                zkteco_device_attend_id = ''

                for line in employee.biometric_device_ids:
                    if line.device_id == biometric_device:
                        zkteco_device_attend_id = line.zkteco_device_attend_id

                ip = biometric_device.zkteco_device_ip_address
                port = biometric_device.port
                password = biometric_device.zkteco_device_pass

                zk = ZK(ip, port, password=password)
                conn = zk.connect()

                if conn:
                    users = zk.get_users()
                    valid = False

                    if users:
                        for user in users:
                            if user.user_id == zkteco_device_attend_id:
                                valid = True

                        if valid:
                            raise ValidationError(f"Already in machine: {biometric_device.name}")
                        else:
                            if zkteco_device_attend_id:
                                zk.set_user(int(zkteco_device_attend_id),
                                            employee.name, 0, '', '',
                                            zkteco_device_attend_id)
                            else:
                                self.update_zkteco_device_emp(biometric_device)
                    else:
                        if zkteco_device_attend_id:
                            zk.set_user(int(zkteco_device_attend_id),
                                        employee.name, 0, '', '',
                                        zkteco_device_attend_id)
                        else:
                            self.update_zkteco_device_emp(biometric_device)

                    zk.disconnect()
                else:
                    raise ValidationError(f"Connection failed for {biometric_device.name}")

    def update_zkteco_device_emp(self, biometric):


        uid_list = []
        user_id_list = []

        ip = biometric.zkteco_device_ip_address
        port = biometric.port
        password = biometric.zkteco_device_pass

        zk = ZK(ip, port, password=password)
        conn = zk.connect()

        if conn:
            zk.disable_device()

            zk.enable_device()

            users = zk.get_users()
            if users:
                for user in users:
                    uid_list.append(user.uid)  # Collect UIDs
                    user_id_list.append(int(user.user_id))  # Collect numeric user IDs

                uid_list.sort()
                user_id_list.sort()

            uid = uid_list[-1] if uid_list else 0
            user_id = user_id_list[-1] if user_id_list else 0

            employee_id = self._context.get('active_id')
            employee = self.env['hr.employee'].search([('id', '=', employee_id)])

            biometric_device = employee.biometric_device_ids.search([
                ('employee_id', '=', employee.id),
                ('device_id', '=', biometric.id)
            ])

            if not biometric_device:
                uid += 1
                user_id += 1

                employee.biometric_device_ids = [(0, 0, {
                    'employee_id': employee.id,
                    'zkteco_device_attend_id': user_id,
                    'device_id': biometric.id,
                })]

                zk.set_user(uid, employee.name, 0, '', '', str(user_id))

            zk.disconnect()

        else:
            raise ValidationError(f"Connection Failed for device: {biometric.name}")

    def action_confirm_biometric_scan(self):


        employee_id = self._context.get('active_id')
        employee = self.env['hr.employee'].search([('id', '=', employee_id)], limit=1)

        for biometric in self.zkteco_device_ids:
            zkteco_device_attend_id = ''

            for attendance_id in employee.biometric_device_ids:
                zkteco_device_attend_id = attendance_id.zkteco_device_attend_id

            ip = biometric.zkteco_device_ip_address
            port = biometric.port
            password = biometric.zkteco_device_pass

            zk = ZK(ip, port, password=password)
            conn = zk.connect()

            if conn:
                users = zk.get_users()
                user_exists = False

                if users:
                    for user in users:
                        if user.user_id == zkteco_device_attend_id:
                            user_exists = True
                            break

                    if user_exists:
                        try:
                            zk.enroll_user(uid=int(zkteco_device_attend_id), user_id=str(zkteco_device_attend_id))
                        except Exception as e:
                            raise UserError(_("An error occurred during fingerprint enrollment: %s") % str(e))

                        raise ValidationError(
                            _("Please place your finger on the biometric device to complete enrollment."))
                    else:
                        raise ValidationError(_("The employee is not registered on the selected biometric device."))
                else:
                    raise ValidationError(_("No user records found on the biometric device."))
            else:
                raise ValidationError(_("Failed to establish a connection with the biometric device."))

    def action_unlink_zkteco_device_employee(self):


        employee_id = self._context.get('active_id')
        employee = self.env['hr.employee'].search([('id', '=', employee_id)], limit=1)

        for biometric in self.zkteco_device_ids:
            if biometric.is_adms:
                zketco_duser_id = self.env['zkteco.attendance.machine'].search([
                    ('employee_id', '=', employee.id),
                    ('device_id', '=', biometric.id),
                ])
                if zketco_duser_id:
                    employee.employee_del_command(biometric)
                else:
                    raise ValidationError(_("The employee is not registered on the selected biometric device."))
            else:
                zkteco_device_attend_id = ''
                for attendance_id in employee.biometric_device_ids:
                    zkteco_device_attend_id = attendance_id.zkteco_device_attend_id

                ip = biometric.zkteco_device_ip_address
                port = biometric.port
                password = biometric.zkteco_device_pass

                zk = ZK(ip, port, password=password)
                conn = zk.connect()

                if conn:
                    users = zk.get_users()
                    user_exists = False

                    if users:
                        for user in users:
                            if user.user_id == zkteco_device_attend_id:
                                user_exists = True
                                break

                        if user_exists:
                            zk.delete_user(user_id=str(zkteco_device_attend_id))

                            biometric_device = employee.biometric_device_ids.search([
                                ('zkteco_device_attend_id', '=', zkteco_device_attend_id),
                                ('device_id', '=', biometric.id)
                            ])
                            biometric_device.unlink()
                        else:
                            raise ValidationError(_("The employee record was not found on the biometric device."))
                    else:
                        raise ValidationError(_("No user records found on the biometric device."))
                else:
                    raise ValidationError(_("Unable to establish a connection with the biometric device."))


class ZKTecoSuccess(models.TransientModel):
    """
    Transient model for displaying a success message to the user.

    This wizard can be triggered after an operation is successfully
    completed (e.g., syncing employees, downloading users, etc.).
    """
    _name = 'zkteco_success'
    _description = 'Success Wizard'



class EmployeeSyncWizard(models.TransientModel):
    """
    Transient model to handle employee synchronization process.

    This wizard is used to initiate and confirm the syncing of
    employees to biometric devices or external systems.
    """
    _name = 'employee.sync.wizard'
    _description = 'Employee Sync Wizard'
