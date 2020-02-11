# -*- coding: utf-8 -*-

import time
from odoo import api, fields, models, _
from odoo.addons import decimal_precision as dp
from odoo.exceptions import UserError
from odoo.exceptions import ValidationError, UserError
from odoo.tools.misc import xlwt
import io
import base64
import operator
import itertools

REPORT_TYPE=[('summery', 'Summery'), ('usage', 'Usage Variance'), ('detailed', 'Detailed')]
class mrp_analysis(models.TransientModel):
    _name = "mrp.analysis"

    type = fields.Selection(REPORT_TYPE,string="Report Type",required=True, )
    mrp_report_file = fields.Binary('MRP Report')
    date_from = fields.Datetime(string="From", required=True, )
    date_to = fields.Datetime(string="To", required=True, )
    categ_id = fields.Many2one(comodel_name="product.category", string="Category",related="product_id.categ_id", required=False, )
    # production_id = fields.Many2one(comodel_name="mrp.production", string="MRP", required=False, )
    production_id = fields.Many2many(comodel_name="mrp.production", relation="name",  string="MRP", )
    product_id = fields.Many2one(comodel_name="product.product", string="Product", required=False, )
    file_name = fields.Char('File Name')
    is_printed = fields.Boolean('Report Printed')

    @api.multi
    def print_other_report(self):
        self.is_printed = False
        for wizard in self:
            return {
                'view_mode': 'form',
                'res_id': wizard.id,
                'res_model': 'mrp.analysis',
                'view_type': 'form',
                'type': 'ir.actions.act_window',
                'context': self.env.context,
                'target': 'new',
            }

    @api.onchange('categ_id')
    def get_productse(self):

        if self.categ_id:
            ids = self.env['product.product'].search([('categ_id', '=', self.categ_id.id)])
            print(ids)

            return {
                'domain': {'product_id': [('id', 'in', ids.ids)], }
            }
        else:
            ids_2 = self.env['product.product'].search([])

            return {
                'domain': {'product_id': [('id', 'in', ids_2.ids)], }}

    @api.multi
    def mrp_analysis_report(self):
        self.is_printed = True


        ctx = dict(self.env.context) or {}
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet("Sheet 1", cell_overwrite_ok=True)
        wb = xlwt.Workbook(style_compression=2)
        column_heading_style = xlwt.easyxf('font:height 300;font:bold True;')
        row = 2
        lines = []
        line_ids = []
        total_f_qty = 0.0
        total_cost = 0.0
        total_per_unit = 0.0
        ####################
        total_s_qty = 0.0
        total_c_qty = 0.0
        total_v_qty = 0.0

        mrp_production = self.env['mrp.production'].search([])


################################################################################################################################################################################
        ##Header Files and Condition as Per Report Type Selected 'Summery Or Usage Or Detailed'''

        if self.type == 'summery':
            report_head = 'Cost Of Products From'+" "+str(self.date_from) +" "+'To'+" "+str(self.date_to)
        elif self.type == 'usage':
            report_head = 'Materials variance per Manufacturing order from'+" "+str(self.date_from)+" "+'To'+" " +str(self.date_to)
        elif self.type == 'detailed':
            report_head = 'Cost sheet of'+" "+str(self.product_id.name)+" "+ 'from'+" "+ str(self.date_from)+ " "+ 'To' +" "+str(self.date_to)

        #summery header
        if self.type == 'summery':
            worksheet.write_merge(0, 2, 2, 7, report_head, xlwt.easyxf(
                'font:height 300; align: vertical center; align: horiz center;pattern: pattern solid, fore_color black; font: color white; font:bold True;' "borders: top thin,bottom thin"))
        #usage header
        elif self.type == 'usage':
            worksheet.write_merge(0, 2, 2, 9, report_head, xlwt.easyxf(
                'font:height 300; align: vertical center; align: horiz center;pattern: pattern solid, fore_color black; font: color white; font:bold True;' "borders: top thin,bottom thin"))

        elif self.type == 'detailed':
            worksheet.write_merge(0, 2, 0, 12, report_head, xlwt.easyxf(
                'font:height 300; align: vertical center; align: horiz center;pattern: pattern solid, fore_color black; font: color white; font:bold True;' "borders: top thin,bottom thin"))

        if self.date_from and self.date_to:
            mrp_production_obj = mrp_production.search(
                [('create_date', '>=', self.date_from), ('create_date', '<=', self.date_to)])
            if self.product_id:
                mrp_production_obj = mrp_production.search(
                    [('create_date', '>=', self.date_from), ('create_date', '<=', self.date_to),
                     ('product_id', '=', self.product_id.id)])
            if self.production_id:
                mrp_production_obj = mrp_production.search(
                    [('create_date', '>=', self.date_from), ('create_date', '<=', self.date_to),
                     ('id', '=', self.production_id.ids)])
        if self.production_id:
            mrp_production_obj = mrp_production.search(
                [('id', '=', self.production_id.ids)])

        if not mrp_production_obj:
            raise ValidationError(_("No Data From This Period"))
        else:
####################################################################################################################################################################
            #When Choose Type """Summary""" Calling This Code and Genrate XLSXS File
####################################################################################################################################################################


            if self.type=='summery':
                for production in mrp_production_obj:

                    worksheet.write(3, 0, _('Product Code'), column_heading_style)
                    worksheet.write(3, 1, _('Name'), column_heading_style)
                    worksheet.write(3, 2, _('Finished Qty'), column_heading_style)
                    worksheet.write(3, 3, _('Total Cost'), column_heading_style)
                    worksheet.write(3, 4, _('Cost Per Unit'), column_heading_style)

                    worksheet.col(0).width = 5000
                    worksheet.col(1).width = 10000
                    worksheet.col(2).width = 5000
                    worksheet.col(3).width = 5000
                    worksheet.col(4).width = 10000
                    worksheet.row(3).height = 500

                    vals = {
                        'product_code': '',
                        'product_name': '',
                        'finished_qty': 0.0,
                        'total_cost': 0.0,
                        'cost_per_unit': 0.0,
                    }

                    for lns in production.finished_move_line_ids:
                        if lns.product_id.id == production.product_id.id:
                            print("LNS", lns.product_id.id)
                            print("production", production.product_id.id)
                            vals = {
                                'product_code': lns.product_id.default_code,
                                'name': lns.product_id.name,
                                'finished_qty': lns.qty_done,
                                'total_cost': production.amount,
                                'cost_per_unit': production.calculate_price,
                            }
                            lines.append(vals)
                            total_f_qty=total_f_qty+lns.qty_done
                            total_cost=total_cost+production.amount
                            total_per_unit=total_per_unit+production.calculate_price
                            row += 2
                            worksheet.write(row, 0, lns.product_id.default_code)
                            worksheet.write(row, 1, lns.product_id.name)
                            worksheet.write(row, 2, lns.qty_done)
                            worksheet.write(row, 3, production.amount)
                            worksheet.write(row, 4, production.calculate_price)
                row += 5
                worksheet.write(row, 0, _('Total Valuation :'),xlwt.easyxf('font:height 200;font:bold True;'))
                worksheet.write(row, 2, (total_f_qty),xlwt.easyxf('font:height 200;font:bold True;'))
                worksheet.write(row, 3, (total_per_unit),xlwt.easyxf('font:height 200;font:bold True;'))
                worksheet.write(row, 4, (total_per_unit),xlwt.easyxf('font:height 200;font:bold True;'))
#######################################################################################################################################################################
                # When Choose Type """Usage""" Calling This Code and Genrate XLSXS File
#######################################################################################################################################################################

            elif self.type=='usage':

                for production in mrp_production_obj:

                    worksheet.write(3, 0, _('MO Number'), column_heading_style)
                    worksheet.write(3, 1, _('Product'), column_heading_style)
                    worksheet.write(3, 2, _('Standard Quantity'), column_heading_style)
                    worksheet.write(3, 3, _('Actual Quantity'), column_heading_style)
                    worksheet.write(3, 4, _('Variance'), column_heading_style)

                    worksheet.col(0).width = 5000
                    worksheet.col(1).width = 10000
                    worksheet.col(2).width = 10000
                    worksheet.col(3).width = 10000
                    worksheet.col(4).width = 5000
                    worksheet.row(3).height = 500

                    vals = {
                        'mo_number': '',
                        'product_id':'',
                        'standard_qty': '',
                        'actual_qty': 0.0,
                        'variance': 0.0,
                    }

                    for lns in production.move_raw_ids:

                            vals = {
                                'mo_number': production.name,
                                'product_id': lns.product_id.name,
                                'standard_qty': lns.product_uom_qty,
                                'actual_qty': lns.quantity_done,
                                'variance':lns.quantity_done -lns.product_uom_qty,
                            }
                            lines.append(vals)
                            total_s_qty = total_s_qty + lns.product_uom_qty
                            total_c_qty = total_c_qty + lns.quantity_done
                            total_v_qty = total_v_qty + lns.quantity_done -lns.product_uom_qty
                            row += 2
                            worksheet.write(row, 0, production.name)
                            worksheet.write(row, 1, lns.product_id.name)
                            worksheet.write(row, 2, lns.product_uom_qty)
                            worksheet.write(row, 3, lns.quantity_done)
                            worksheet.write(row, 4, lns.quantity_done -lns.product_uom_qty)

                    row += 3
                    worksheet.write(row, 0, _('Total:%s')%(production.name),xlwt.easyxf('font:height 200;font:bold True;'))
                    worksheet.write(row, 2, (total_s_qty),xlwt.easyxf('font:height 200;font:bold True;'))
                    worksheet.write(row, 3, (total_c_qty),xlwt.easyxf('font:height 200;font:bold True;'))
                    worksheet.write(row, 4, (total_v_qty),xlwt.easyxf('font:height 200;font:bold True;'))
########################################################################################################################################################################
            #When Choose Type """Detiled""" Calling This Code and Genrate XLSXS File
########################################################################################################################################################################

            elif self.type == 'detailed':
                for production in mrp_production_obj:

                    worksheet.write(3, 0, _('Item code'), column_heading_style)
                    worksheet.write(3, 1, _('Item name'), column_heading_style)
                    worksheet.write(3, 2, _('Standard Quantity'), column_heading_style)
                    worksheet.write(3, 3, _('Consumed Quantity'), column_heading_style)
                    worksheet.write(3, 4, _('Quantity Variance'), column_heading_style)
                    worksheet.write(3, 5, _('Av.Price'), column_heading_style)
                    worksheet.write(3, 6, _('Variance value'), column_heading_style)
                    worksheet.write(3, 7, _('Total Cost'), column_heading_style)

                    worksheet.col(0).width = 5000
                    worksheet.col(1).width = 10000
                    worksheet.col(2).width = 10000
                    worksheet.col(3).width = 10000
                    worksheet.col(4).width = 10000
                    worksheet.col(5).width = 5000
                    worksheet.col(6).width = 10000
                    worksheet.col(7).width = 10000
                    worksheet.row(3).height = 500

                    vals = {
                        'product_code': '',
                        'product_name': '',
                        'standard_qty': 0.0,
                        'consumed_qty': 0.0,
                        'qty_varince':0.0,
                        'avrage_price': 0.0,
                        'variance_value': 0.0,
                        'total_cost': 0.0,
                    }

                    for lns in production.finished_move_line_ids:
                        if lns.product_id.id == production.product_id.id:
                            vals = {
                                'product_code': lns.product_id.default_code,
                                'name': lns.product_id.name,
                                'standard_qty': lns.product_uom_qty,
                                'consumed_qty': lns.qty_done,
                                'qty_varince': lns.product_uom_qty-lns.qty_done,
                                'avrage_price': production.calculate_price,
                                'variance_value': (lns.product_uom_qty-lns.qty_done)*production.calculate_price,
                                'total_cost':lns.qty_done*production.calculate_price,
                            }
                            lines.append(vals)
                            # total_f_qty = total_f_qty + lns.qty_done
                            # total_cost = total_cost + production.amount
                            # total_per_unit = total_per_unit + production.calculate_price
                            row +=2
                            worksheet.write(row, 0, lns.product_id.default_code)
                            worksheet.write(row, 1, lns.product_id.name)
                            worksheet.write(row, 2, lns.product_uom_qty)
                            worksheet.write(row, 3, lns.qty_done)
                            worksheet.write(row, 4, lns.product_uom_qty-lns.qty_done)
                            worksheet.write(row, 5, production.calculate_price)
                            worksheet.write(row, 6, (lns.product_uom_qty-lns.qty_done)*production.calculate_price)
                            worksheet.write(row, 7, lns.qty_done*production.calculate_price)
                row += 10
                worksheet.row(row).height = 500

                worksheet.write(row, 0, _('By Product'),  xlwt.easyxf('font:height 300;font:bold True;pattern:fore_color black;'))
                row += 1
                worksheet.write(row, 0, _('Item code'), column_heading_style)
                worksheet.write(row, 1, _('Item name'), column_heading_style)
                worksheet.write(row, 2, _('Standard Quantity'), column_heading_style)
                worksheet.write(row, 3, _('Consumed Quantity'), column_heading_style)
                worksheet.write(row, 4, _('Quantity Variance'), column_heading_style)
                worksheet.write(row, 5, _('Av.Price'), column_heading_style)
                worksheet.write(row, 6, _('Variance value'), column_heading_style)
                worksheet.write(row, 7, _('Total Cost'), column_heading_style)
                worksheet.row(row).height = 500

                for lns in production.finished_move_line_ids:
                    if lns.product_id.id != production.product_id.id:
                        vals = {
                            'product_code': lns.product_id.default_code,
                            'name': lns.product_id.name,
                            'standard_qty': lns.product_uom_qty,
                            'consumed_qty': lns.qty_done,
                            'qty_varince': lns.product_uom_qty - lns.qty_done,
                            'avrage_price': production.calculate_price,
                            'variance_value': (lns.product_uom_qty - lns.qty_done) * production.calculate_price,
                            'total_cost': lns.qty_done * production.calculate_price,
                        }
                        lines.append(vals)
                        # total_f_qty = total_f_qty + lns.qty_done
                        # total_cost = total_cost + production.amount
                        # total_per_unit = total_per_unit + production.calculate_price
                        row += 2
                        worksheet.write(row, 0, lns.product_id.default_code)
                        worksheet.write(row, 1, lns.product_id.name)
                        worksheet.write(row, 2, lns.product_uom_qty)
                        worksheet.write(row, 3, lns.qty_done)
                        worksheet.write(row, 4, lns.product_uom_qty - lns.qty_done)
                        worksheet.write(row, 5, production.calculate_price)
                        worksheet.write(row, 6, (lns.product_uom_qty - lns.qty_done) * production.calculate_price)
                        worksheet.write(row, 7, lns.qty_done * production.calculate_price)


                # worksheet.write(row, 0, _('Total Valuation :'))
                # worksheet.write(row, 2, total_f_qty)
                # worksheet.write(row, 3, total_per_unit)
                # worksheet.write(row, 4, total_per_unit)

########################################################################################################################################################################

        for wizard in self:
            # report_head = 'Warehouse Stock Report'
            fp = io.BytesIO()
            workbook.save(fp)
            excel_file = base64.encodestring(fp.getvalue())
            wizard.mrp_report_file = excel_file
            if self.type=='summery':
                wizard.file_name = 'MRP Summery Report.xls'
            elif self.type=='usage':
                wizard.file_name = 'MRP Usage Variance Report.xls'
            elif self.type == 'detailed':
                wizard.file_name = 'MRP Detailed Report.xls'
            fp.close()
            return {
                'view_mode': 'form',
                'res_id': wizard.id,
                'res_model': 'mrp.analysis',
                'view_type': 'form',
                'type': 'ir.actions.act_window',
                'context': self.env.context,
                'target': 'new',
            }
