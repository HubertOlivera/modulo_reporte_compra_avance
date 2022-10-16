# -*- coding: utf-8 -*-
from odoo import models, fields, api, _
from odoo.exceptions import UserError
import base64
from odoo.tools.float_utils import float_round

class PurchaseOrderReport(models.TransientModel):
	_name = 'purchase.order.report'
	_description = 'Purchase Order Report'

	date_from = fields.Date(string='Fecha Desde', required=True)
	date_to = fields.Date(string='Fecha Hasta', default=fields.Date.today(), required=True)

	def get_purchase_report(self):
		import io
		from xlsxwriter.workbook import Workbook
		MainParameter = self.env['main.parameter'].get_main_parameter()
		ReportBase = self.env['report.base']
		if not MainParameter.dir_create_file:
			raise UserError('Falta configurar un directorio de descargas en Parametros Principales')
		route = MainParameter.dir_create_file
		workbook = Workbook(route + 'Purchase Order.xlsx')
		workbook, formats = ReportBase.get_formats(workbook)

		import importlib
		import sys
		importlib.reload(sys)

		worksheet = workbook.add_worksheet('ORDENES DE COMPRA')
		worksheet.set_tab_color('blue')
		HEADERS = [
			'NOMBRE DE LA O/C',
			'EMPRESA',
			'FECHA',
			'FECHA DE ENVIO A CONTABILIDAD',
			'MONEDA',
			'TIPO DE CAMBIO',
			'IMPORTE EN MONEDA ORIGINAL',
			'TOTAL',
			'DESCRIPCION',
		]
		worksheet = ReportBase.get_headers(worksheet, HEADERS, 0, 0, formats['boldbord'])

		Orders = self.env['purchase.order'].search([
			('date_order', '>=', self.date_from),
			('date_order', '<=', self.date_to),
			('state', '=', 'purchase'),
		])
		x = 1
		total = 0
		for order in Orders:
			worksheet.write(x, 0, f'Orden de Compra {order.name}', formats['especial1'])
			worksheet.write(x, 1, order.partner_id.name, formats['especial1'])
			worksheet.write(x, 2, order.date_order, formats['reverse_dateformat'])
			worksheet.write(x, 3, order.date_approve, formats['reverse_dateformat'])
			worksheet.write(x, 4, order.currency_id.name, formats['especial1'])
			if order.currency_id == self.env.company.currency_id:
				currency_rate = 1.0
			else:
				Rate = self.env['res.currency.rate'].search([('name', '=', order.date_approve)])
				currency_rate = Rate.sale_type if Rate else 1.0
			amount_pen = float_round(order.amount_total * currency_rate, 2)
			worksheet.write(x, 5, currency_rate, formats['numberdos'])
			worksheet.write(x, 6, order.amount_total, formats['especial1'])
			worksheet.write(x, 7, amount_pen, formats['numberdos'])
			worksheet.write(x, 8, order.glosa if order.glosa else '', formats['especial1'])

			x += 1
			total += amount_pen
		worksheet.write(x, 6, 'TOTAL', formats['boldbord'])
		worksheet.write(x, 7, total, formats['numbertotal'])

		widths = [18, 36, 12, 12, 12, 12, 10, 12, 36]
		worksheet = ReportBase.resize_cells(worksheet, widths)

		workbook.close()

		f = open(route + 'Purchase Order.xlsx', 'rb')
		return self.env['popup.it'].get_file('Purchase Order.xlsx', base64.encodestring(b''.join(f.readlines())))
