# -*- coding: utf-8 -*-
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2024 DOTSPRIME SYSTEM LLP
#    Email : sales@dotsprime.com / dotsprime@gmail.com
########################################################

from odoo import http, SUPERUSER_ID
from odoo.http import request
from werkzeug.wrappers import Response
import time
from datetime import datetime


class ZKTecoController(http.Controller):
    """
    Controller for ZKTeco biometric device integration.
    Handles communication between device and Odoo:
    - Device connectivity and configuration
    - Receiving logs (attendance & operation)
    - Command dispatch and response
    """
    def generate_zkteco_op_bid_logs(self, raw_data, device_id, opStamp):
        dict = {
            'device_id': device_id.id,
            'opStamp': opStamp,
            'log_text': raw_data

        }
        request.env['device.operation.stamplogs'].sudo().create(dict)

    def generate_zkteco_slogs(self, raw_data, device_id, Stamp):
        dict = {
            'device_id': device_id.id,
            'stamp': Stamp,
            'log_text': raw_data,

        }
        request.env['device.stamp.logs'].sudo().create(dict)

    @http.route('/iclock/cdata', type='http', auth='public', methods=['GET'])
    def zkteco_cdata(self, **kwargs):
        """
        Handle ZKTeco device connectivity and provide configuration response.

        This endpoint is called by the ZKTeco biometric device when it attempts to connect
        to the server. It verifies the device by its serial number and returns configuration
        details such as transfer timings, intervals, and last processed stamps (Stamp and OpStamp).

        **Request Parameters**:
            SN (str): Serial number of the device
            options (str, optional): Device options
            pushver (str, optional): Device push version
            language (str, optional): Device language

        **Response**:
            str: A configuration string containing:
                 - Last attendance stamp (Stamp)
                 - Last operation stamp (OpStamp)
                 - Error and delay configurations
                 - Transfer time slots and intervals
                 - Flags and encryption settings
            If the device is not found in the system, returns an HTTP 405 response.
        """

        sn = kwargs.get('SN')  # Device Serial Number
        options = kwargs.get('options')
        pushver = kwargs.get('pushver')
        language = kwargs.get('language')  # Device language setting

        device_id = request.env['zkteco.device.setting'].sudo().search([
            ('serial_number', '=', sn)
        ])

        if device_id:
            device_id.sudo().state = 'connected'
            now = datetime.now()
            fixed_time = "00:00"
            current_time = now.strftime("%H:%M")
            formatted_time = f"{fixed_time};{current_time}"

            operation_log = request.env['device.operation.stamplogs'].sudo().search([])
            attendance_log_ids = request.env['device.stamp.logs'].sudo().search([])

            opStamp = operation_log.sorted('opStamp')[-1].opStamp if operation_log else 0
            stamp = attendance_log_ids.sorted('stamp')[-1].stamp if attendance_log_ids else 0

            response = (
                f"GET OPTION FROM: {sn}\n"
                f"Stamp={stamp}\n"
                f"OpStamp={opStamp}\n"
                f"ErrorDelay={device_id.error_delay}\n"
                f"Delay={device_id.delay}\n"
                f"TransTimes={formatted_time}\n"
                f"TransInterval={device_id.device_t_interval}\n"
                f"TransFlag=1101111000\n"
                f"Realtime=1\n"
                f"Encrypt=0\n"
            )

            return response

        return Response("No matching device found. Ensure the device is properly registered.", 405)

    @http.route('/iclock/cdata', type='http', auth='none', methods=['POST'], csrf=False)
    def fetch_zketco_bid_datas(self, **kwargs):
        """
        Handle POST requests from ZKTeco biometric devices for log synchronization.

        This endpoint receives logs from the ZKTeco device and processes them into Odoo.
        The logs may include:
            - Attendance logs (ATTLOG)
            - Operation logs (OPERLOG), which may include user and fingerprint data.

        **Request Parameters**:
            SN (str): Serial number of the device.
            Stamp (str, optional): Last processed attendance log stamp.
            OpStamp (str, optional): Last processed operation log stamp.
            table (str): Type of log data being sent ("ATTLOG" or "OPERLOG").

        **Behavior**:
            - Decodes raw POST data sent by the device.
            - Identifies the device in the system.
            - Processes logs based on the table type:
                * OPERLOG → Operation logs, user data, fingerprints.
                * ATTLOG → Attendance logs.
            - Creates records in relevant models.

        **Response**:
            str: "OK" if processed successfully.
        """

        serial_number = kwargs.get('SN')  # Device Serial Number
        Stamp = kwargs.get('Stamp')
        OpStamp = kwargs.get('OpStamp')
        table = kwargs.get('table')  # Log type (OPERLOG or ATTLOG)

        stp_value = Stamp if Stamp else OpStamp

        # base_data = http.request.httprequest.data.decode('utf-8')
        raw_data = http.request.httprequest.data
        try:
            base_data = raw_data.decode('utf-8')
        except UnicodeDecodeError as e:
            # fallback nếu có byte không hợp lệ
            import logging
            _logger = logging.getLogger(__name__)
            _logger.warning(f"⚠️ ZKTeco log decode error: {e}. Fallback to latin-1 decoding.")
            base_data = raw_data.decode('latin-1', errors='ignore')
        
        # Custimized by Tunn
        # ⚙️ Dùng môi trường admin
        env = request.env(user=SUPERUSER_ID)
        device_id = env['zkteco.device.setting'].search([('serial_number', '=', serial_number)])

        # device_id = request.env['zkteco.device.setting'].sudo().search([
        #     ('serial_number', '=', serial_number)
        # ])

        if device_id:
            if serial_number and table == "OPERLOG":
                self.generate_zkteco_op_bid_logs(base_data, device_id, stp_value)

                for line in base_data.strip().split('\n'):
                    if line.startswith("OPLOG"):
                        values = line.split()
                        try:
                            device_id.create_oplog(values, stp_value)
                        except Exception as e:
                            print(f"Error processing OPLOG: {e}")
                    elif line.startswith("FP"):
                        values = line.split()
                        device_id.action_create_device_user_fingerprint(values)
                    elif line.startswith("USER"):
                        values = line.split()
                        device_id.action_create_employee_device_user(values)

            if serial_number and table == "ATTLOG":
                self.generate_zkteco_slogs(base_data, device_id, stp_value)

                device_id.action_create_device_zkteco_logs(base_data)
        return Response("OK", 200)
    @http.route('/iclock/getrequest', type='http', auth='public', methods=['GET'], csrf=False)
    def get_request(self, **kwargs):
        """
        Handle GET requests from ZKTeco devices to retrieve pending commands.

        This endpoint is called by ZKTeco devices to check if there are any pending
        commands from the Odoo server. Commands are typically created when new users,
        fingerprints, or configurations need to be synced to the device.

        **Request Parameters**:
            SN (str): Serial number of the ZKTeco device.

        **Behavior**:
            - Identifies the device using its serial number.
            - Fetches the next pending command for the device from Odoo.
            - Returns the command if available; otherwise, responds with "OK".

        **Response**:
            str: A pending command for the device, or "OK" if no commands exist.
        """
        device_sn = kwargs.get('SN')
        device_id = request.env['zkteco.device.setting'].sudo().search([
            ('serial_number', '=', device_sn)
        ])
        command = device_id.action_create_zkteco_device_user_commands()

        return Response(command if command else "OK", 200)

    @http.route('/iclock/devicecmd', type='http', auth='public', methods=['POST'], csrf=False)
    def zkteco_bid_operation_cmd(self, **kwargs):
        """
        Handle acknowledgment of executed commands from ZKTeco devices.

        This endpoint receives POST requests from ZKTeco devices after they execute
        commands sent by Odoo (e.g., add user, delete user). It parses the acknowledgment
        data, identifies the executed command, and updates its status in Odoo.

        **Request Parameters (from URL)**:
            - SN (str): Serial number of the ZKTeco device.

        **Request Body**:
            - Raw data containing acknowledgment details in key-value pairs
              (e.g., "CMD=DATA&ID=1").

        **Behavior**:
            - Decode the incoming raw data.
            - Extract acknowledgment parameters.
            - Identify the executed command (only if CMD is 'DATA' or 'CHECK').
            - Update the command status in Odoo.

        **Response**:
            str: "OK" to confirm successful processing.
        """

        base_data = http.request.httprequest.data.decode('utf-8')

        serial_number = kwargs.get('SN')

        device_id = request.env['zkteco.device.setting'].sudo().search([
            ('serial_number', '=', serial_number)
        ])

        for line in base_data.split('\n'):
            if not line.strip():
                continue

            parameters = line.split('&')
            parameters_dictionary = {}
            for param in parameters:
                if param:
                    key, value = param.split("=")
                    parameters_dictionary[key] = value

            if parameters_dictionary.get("CMD") in ["DATA", "CHECK"]:
                command_id = parameters_dictionary.get("ID")
                device_id.action_check_zkteco_device_command_revert_res(command_id)
        return Response("OK", 200)