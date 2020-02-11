# -*- coding: utf-8 -*-
{
    'name': "MRP Analysis Excel Report",

    'summary': """
        you Can From This Model Genrate Dynamic Reports From Manufacturing Orders Consumed martial and Finished , Scrap order By Product
          Contain From To Types Summery and Usage Variance and Detailed """,

    'author': "El-Sayed Iraky",
    'category': 'MRP',
    'version': '0.1',
    'depends': ['base','mrp','stock','report_xlsx'],
    'data': [
        # 'security/ir.model.access.csv',
        'views/mrp_analysis.xml',
    ],

}