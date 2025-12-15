{
    'name': 'EAUT ZKTeco Integration',
    'version': '19.0.1.0.0',
    'category': 'Human Resources',
    'summary': 'Automate attendance by integrating ZKTeco biometric devices with Odoo.',
    'author': 'Dotsprime System',
    'website': 'https://dotsprime.com/',
    'license': 'AGPL-3',

    'depends': [
        'base',
        'hr',
        'hr_attendance',
    ],

    'data': [
        'security/ir.model.access.csv',
        'security/security.xml',
        'views/dashboard_dashboard.xml',
        'demo/dashboard_dashboard_demo.xml',
        'data/ir_cron.xml',
        'views/zkteco_device_settings_views.xml',
        'views/zkteco_device_logs.xml',
        'wizard/zkteco_device_attendance_create.xml',
        'wizard/zkteco_attendance_device_view.xml',
        'wizard/zkteco_device_attendance_report_view.xml',
        'wizard/employee_leave_assign_wizard.xml',
        'wizard/attendance_reports.xml',
        'views/views_inherit.xml',
        'views/attendance_state_views.xml',
        'views/zkteco_device_fingerprints.xml',
        'views/device_views.xml',
        'views/hr_attendance_view.xml',
        'views/hr_employee_view.xml',
        'views/resource_calendar_attendance_view.xml',
        'views/device_user_views.xml',
        'views/menus.xml',
    ],

    'assets': {
        'web.assets_backend': [
            'dps_zkteco_biometric_integration/static/src/scss/zkteco_dashboard.scss',
        ],
    },

    'installable': True,
    'application': True,
}
