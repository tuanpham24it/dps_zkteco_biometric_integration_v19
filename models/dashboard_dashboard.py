from odoo import api,fields,models,_
import time
from dateutil.relativedelta import relativedelta
from odoo.exceptions import UserError
from datetime import datetime,timedelta
from babel.dates import format_datetime,format_date
from odoo.tools import DEFAULT_SERVER_DATE_FORMAT as DF,DEFAULT_SERVER_DATETIME_FORMAT as DTF,format_datetime as tool_format_datetime
from odoo.release import version
import json

DASHBOARD_FIELDS = ['total_attedance_logs','my_total_attedance_logs','total_attedance_state','my_total_attedance_state_color','total_device','my_total_device_color','total_employee','my_total_employee_color','total_absent','my_total_absent_color','total_present','my_total_present_color','total_late','my_total_late_color','total_early_leave','my_total_early_leave_color','present_employee_data','absent_employee_data','late_employee_data','early_leave_employee_data']

class DashboardDashboard(models.Model):
	_name = 'dashboard.dashboard'
	_description = "Dashboard Dashboard"


	@property
	def SELF_READABLE_FIELDS(self):
		return super().SELF_READABLE_FIELDS + DASHBOARD_FIELDS

	@property
	def SELF_WRITEABLE_FIELDS(self):
		return super().SELF_WRITEABLE_FIELDS + DASHBOARD_FIELDS
		
	dashboard_data_filter = fields.Selection([
		('today','Today'),
		('week','This Week'),
		('month','This Month'),
		('year','This Year'),
		('all','All'),
	],string="Dashboard Filter",default='today')


	total_attedance_logs = fields.Integer(string="Total attedance logs",compute="_compute_total_attendance_logs")
	my_total_attedance_logs = fields.Char(string='My Total attedance logs',default="#883F73")
	total_attedance_state = fields.Integer(string="Total attedance state",compute="_compute_total_attendance_state")
	my_total_attedance_state_color = fields.Char(string='My Total attedance state',default="#883F73")
	total_device = fields.Integer(string="Total Device",compute="_compute_total_device")
	my_total_device_color = fields.Char(string='My Total Device',default="#883F73")
	total_employee = fields.Integer(string="Total employee",compute="_compute_total_employee")
	my_total_employee_color = fields.Char(string='My Total employee',default="#883F73")

	total_absent = fields.Integer(string="Total Absent Employee",compute="_compute_total_absent")
	my_total_absent_color = fields.Char(string='My Total absent color',default="#883F73")

	total_present = fields.Integer(string="Total present Employee",compute="_compute_total_present")
	my_total_present_color = fields.Char(string='My Total present color',default="#883F73")

	total_late = fields.Integer(string="Total Late Arrival", compute="_compute_total_late")
	my_total_late_color = fields.Char(string='My Total Late Color', default="#883F73")

	total_early_leave = fields.Integer(string="Total Early Leaving", compute="_compute_total_early_leave")
	my_total_early_leave_color = fields.Char(string="My Total Early Leave Color", default="#883F73")

	present_employee_data = fields.Text(string="PResent Employee",compute="_compute_present_employee")
	absent_employee_data = fields.Text(string="Absent Employee",compute="_compute_absent_employee")
	late_employee_data = fields.Text(string="Absent Employee",compute="_compute_late_employee")
	early_leave_employee_data = fields.Text(string="Early Leave Employees", compute="_compute_early_leave_employee")

	def today_data(self):
		self.dashboard_data_filter = 'today'

	def week_data(self):
		self.dashboard_data_filter = 'week'

	def month_data(self):
		self.dashboard_data_filter = 'month'

	def year_data(self):
		self.dashboard_data_filter = 'year'

	def all_data(self):
		self.dashboard_data_filter = 'all'

	def get_filter(self,field_name):
		for rec in self:
			domain = []
			if rec.dashboard_data_filter=='today':
				domain = [(field_name,'>=',time.strftime('%Y-%m-%d 00:00:00')),(field_name,'<=',time.strftime('%Y-%m-%d 23:59:59'))]
			if rec.dashboard_data_filter=='week':
				domain = [(field_name,'>=',(fields.Datetime.today() + relativedelta(weeks=-1,days=1,weekday=0)).strftime('%Y-%m-%d')),(field_name,'<=',(fields.Datetime.today() + relativedelta(weekday=6)).strftime('%Y-%m-%d'))]
			if rec.dashboard_data_filter=='month':
				domain = [(field_name,'<',(fields.Datetime.today()+relativedelta(months=1)).strftime('%Y-%m-01')),(field_name,'>=',time.strftime('%Y-%m-01'))]
		return domain

	@api.depends('dashboard_data_filter')
	def _compute_total_attendance_logs(self):
		for rec in self:
			domain = rec.get_filter('create_date')
			rec.total_attedance_logs = self.env['zkteco.device.logs'].search_count(domain)

	@api.depends('dashboard_data_filter')
	def _compute_total_attendance_state(self):
		for rec in self:
			domain = rec.get_filter('create_date')
			rec.total_attedance_state = self.env['zkteco.device.states'].search_count(domain)

	@api.depends('dashboard_data_filter')
	def _compute_total_device(self):
		for rec in self:
			domain = rec.get_filter('create_date')
			rec.total_device = self.env['zkteco.device.setting'].search_count(domain)


	@api.depends('dashboard_data_filter')
	def _compute_total_employee(self):
		for rec in self:
			domain = rec.get_filter('create_date')
			rec.total_employee = self.env['hr.employee'].search_count(domain)


	@api.depends('dashboard_data_filter')
	def _compute_total_absent(self):
		for rec in self:
			employees = self.env['hr.employee'].search(rec.get_filter('create_date'))
			absent_count = 0
			for emp in employees:
				if emp.hr_presence_state == 'absent':
					absent_count += 1
			rec.total_absent = absent_count


	@api.depends('dashboard_data_filter')
	def _compute_total_present(self):
		for rec in self:
			employees = self.env['hr.employee'].search(rec.get_filter('create_date'))
			present_count = 0
			for emp in employees:
				if emp.hr_presence_state == 'present':
					present_count += 1
			rec.total_present = present_count
	

	@api.depends('dashboard_data_filter')
	def _compute_total_late(self):
		for rec in self:
			domain = rec.get_filter('check_in')
			today_date = fields.Date.context_today(self)
			today_9am = fields.Datetime.to_datetime(f"{today_date} 09:00:00")
			attendances = self.env['hr.attendance'].search(domain)
			late_employees = set()
			for att in attendances:
				if att.check_in and att.check_in > today_9am:
					late_employees.add(att.employee_id.id)

			rec.total_late = len(late_employees)


	@api.depends('dashboard_data_filter')
	def _compute_total_early_leave(self):
		for rec in self:
			domain = rec.get_filter('check_out')
			today_date = fields.Date.context_today(self)
			today_7pm = fields.Datetime.to_datetime(f"{today_date} 19:00:00")
			attendances = self.env['hr.attendance'].search(domain)
			early_employee_ids = set()
			for att in attendances:
				if att.check_out and att.check_out < today_7pm:
					early_employee_ids.add(att.employee_id.id)
			rec.total_early_leave = len(early_employee_ids)

	def open_late(self):
		domain = self.get_filter('check_in')
		today_date = fields.Date.context_today(self)
		today_9am = fields.Datetime.to_datetime(f"{today_date} 09:00:00")
		attendances = self.env['hr.attendance'].search(domain)
		late_employee_ids = []
		for att in attendances:
			if att.check_in and att.check_in > today_9am:
				if att.employee_id.id not in late_employee_ids:
					late_employee_ids.append(att.employee_id.id)
		action = self.env["ir.actions.actions"]._for_xml_id("hr.open_view_employee_list_my")
		action['domain'] = [('id', 'in', late_employee_ids)]
		return action



	def open_attendance_log(self):
		action = self.env["ir.actions.actions"]._for_xml_id("dps_zkteco_biometric_integration.action_zkteco_device_attendance_logs")
		action['domain'] = self.get_filter('create_date')
		return action

	def open_attendance_state(self):
		action = self.env["ir.actions.actions"]._for_xml_id("dps_zkteco_biometric_integration.action_zkteco_device_attendance_states")
		action['domain'] = self.get_filter('create_date')
		return action

	def open_device(self):
		action = self.env["ir.actions.actions"]._for_xml_id("dps_zkteco_biometric_integration.action_zkteco_device_settings_view")
		action['domain'] = self.get_filter('create_date')
		return action

	def open_employee(self):
		action = self.env["ir.actions.actions"]._for_xml_id("hr.open_view_employee_list_my")
		action['domain'] = self.get_filter('create_date')
		return action


	def open_absent(self):
		employees = self.env['hr.employee'].search(self.get_filter('create_date'))
		absent_employee_ids = []
		for emp in employees:
			if emp.hr_presence_state == 'absent':
				absent_employee_ids.append(emp.id)
		action = self.env["ir.actions.actions"]._for_xml_id("hr.open_view_employee_list_my")
		action['domain'] = [('id', 'in', absent_employee_ids)]
		return action

	def open_present(self):
		employees = self.env['hr.employee'].search(self.get_filter('create_date'))
		present_employee_ids = []
		for emp in employees:
			if emp.hr_presence_state == 'present':
				present_employee_ids.append(emp.id)
		action = self.env["ir.actions.actions"]._for_xml_id("hr.open_view_employee_list_my")
		action['domain'] = [('id', 'in', present_employee_ids)]
		return action
	

	def open_early_leave(self):
		domain = self.get_filter('check_out')

		today_date = fields.Date.context_today(self)
		today_7pm = fields.Datetime.to_datetime(f"{today_date} 19:00:00")

		attendances = self.env['hr.attendance'].search(domain)

		early_employee_ids = []
		for att in attendances:
			if att.check_out and att.check_out < today_7pm:
				if att.employee_id.id not in early_employee_ids:
					early_employee_ids.append(att.employee_id.id)

		action = self.env["ir.actions.actions"]._for_xml_id("hr.open_view_employee_list_my")
		action['domain'] = [('id', 'in', early_employee_ids)]
		return action

	
	@api.depends('dashboard_data_filter')
	def _compute_present_employee(self):
		for rec in self:
			employees = self.env['hr.employee'].search(rec.get_filter('create_date'))

			present_employee_ids = []
			employee_data = []
			tzinfo = self.env.context.get('tz') or self.env.user.tz or 'UTC'
			locale = self.env.context.get('lang') or self.env.user.lang or 'en_US'

			for emp in employees:
				if emp.hr_presence_state == 'present':
					present_employee_ids.append(emp.id)

			rec.total_present = len(present_employee_ids)

			for emp in self.env['hr.employee'].browse(present_employee_ids[:20]):
				last_attendance = emp.attendance_ids.sorted(key=lambda r: r.check_in, reverse=True)
				last_check_in = last_attendance[0].check_in if last_attendance else ''
				employee_data.append({
					'id': emp.id,
					'name': emp.name,
					'department': emp.department_id.name if emp.department_id else '',
					'job': emp.job_id.name if emp.job_id else '',
					'check_in': tool_format_datetime(self.env, last_check_in) if last_check_in else '',
				})

			rec.present_employee_data = json.dumps(employee_data)

	@api.depends('dashboard_data_filter')
	def _compute_absent_employee(self):
		for rec in self:
			employees = self.env['hr.employee'].search(rec.get_filter('create_date'))
			absent_employee_ids = []
			employee_data = []
			tzinfo = self.env.context.get('tz') or self.env.user.tz or 'UTC'
			locale = self.env.context.get('lang') or self.env.user.lang or 'en_US'
			for emp in employees:
				if emp.hr_presence_state == 'absent':
					absent_employee_ids.append(emp.id)
			rec.total_absent = len(absent_employee_ids)
			for emp in self.env['hr.employee'].browse(absent_employee_ids[:20]):
				last_attendance = emp.attendance_ids.sorted(key=lambda r: r.check_in, reverse=True)
				last_check_in = last_attendance[0].check_in if last_attendance else ''
				employee_data.append({
					'id': emp.id,
					'name': emp.name,
					'department': emp.department_id.name if emp.department_id else '',
					'job': emp.job_id.name if emp.job_id else '',
					'last_check_in': tool_format_datetime(self.env, last_check_in) if last_check_in else '',
				})
			rec.absent_employee_data = json.dumps(employee_data)



	@api.depends('dashboard_data_filter')
	def _compute_late_employee(self):
		for rec in self:
			domain = rec.get_filter('check_in')
			today_date = fields.Date.context_today(self)
			today_9am = fields.Datetime.to_datetime(f"{today_date} 09:00:00")
			attendances = self.env['hr.attendance'].search(domain)
			late_employee_ids = set()
			for att in attendances:
				if att.check_in and att.check_in > today_9am:
					late_employee_ids.add(att.employee_id.id)
			rec.total_late = len(late_employee_ids)

			employee_data = []
			for emp in self.env['hr.employee'].browse(list(late_employee_ids)[:20]):
				last_attendance = emp.attendance_ids.sorted(key=lambda r: r.check_in, reverse=True)
				last_check_in = last_attendance[0].check_in if last_attendance else ''
				employee_data.append({
					'id': emp.id,
					'name': emp.name,
					'department': emp.department_id.name if emp.department_id else '',
					'job': emp.job_id.name if emp.job_id else '',
					'last_check_in': tool_format_datetime(self.env, last_check_in) if last_check_in else '',
				})

			rec.late_employee_data = json.dumps(employee_data or [])

	@api.depends('dashboard_data_filter')
	def _compute_early_leave_employee(self):
		for rec in self:
			domain = rec.get_filter('check_out')
			today_date = fields.Date.context_today(self)
			today_7pm = fields.Datetime.to_datetime(f"{today_date} 19:00:00")
			
			attendances = self.env['hr.attendance'].search(domain)
			early_employee_ids = set()
			
			for att in attendances:
				if att.check_out and att.check_out < today_7pm:
					early_employee_ids.add(att.employee_id.id)
			
			rec.total_early_leave = len(early_employee_ids)
			employee_data = []
			for emp in self.env['hr.employee'].browse(list(early_employee_ids)[:20]):
				last_attendance = emp.attendance_ids.sorted(key=lambda r: r.check_out, reverse=True)
				last_check_out = last_attendance[0].check_out if last_attendance else ''
				employee_data.append({
					'id': emp.id,
					'name': emp.name,
					'department': emp.department_id.name if emp.department_id else '',
					'job': emp.job_id.name if emp.job_id else '',
					'last_check_out': tool_format_datetime(self.env, last_check_out) if last_check_out else '',
				})
			rec.early_leave_employee_data = json.dumps(employee_data or [])

	

	def main_open_dashboard_action(self):
		method = self._context.get('main_action')
		if not method:
			raise UserError("No action Defined to call.")
		result = getattr(self,method)()
		return result


