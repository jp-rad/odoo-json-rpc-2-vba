from odoo import models, fields, api

class VbaJsonSample(models.Model):
    _name = 'vba_json.sample'
    _description = 'VBA Sample Model with All Odoo Data Types'

    # Reserved fields
    name = fields.Char(string='Name', required=True)

    # Basic fields
    boolean_field = fields.Boolean(string='Boolean', default=True)
    char_field = fields.Char(string='Char', default='Default Char')
    float_field = fields.Float(string='Float', digits=(12, 2), default=123.45)
    integer_field = fields.Integer(string='Integer', default=10)

    # Advanced fields
    html_field = fields.Html(string='Html Field')
    currency_id = fields.Many2one('res.currency', string='Currency', default=lambda self: self.env.company.currency_id)
    monetary_field = fields.Monetary(string='Monetary', currency_field='currency_id', default=1000.0)
    selection_field = fields.Selection([
        ('draft', 'Draft'),
        ('in_progress', 'In Progress'),
        ('done', 'Done')
    ], string='Selection', default='draft')
    text_field = fields.Text(string='Text (Multi-line)')

    # Date(time) fields
    date_field = fields.Date(string='Date', default=fields.Date.context_today)
    datetime_field = fields.Datetime(string='Datetime', default=fields.Datetime.now)

    # Relational fields
    many2one_field = fields.Many2one('res.partner', string='Many2one (Partner)')
    one2many_field = fields.One2many('vba_json.sample.line', 'vba_json_sample_id', string='One2many Lines')
    many2many_field = fields.Many2many('res.partner.category', string='Many2many (Tags)')

    # Pseudo-relational fields
    reference_field = fields.Reference([
        ('res.partner', 'Partner'),
        ('res.users', 'User')
    ], string='Reference')
    ref_model = fields.Char(string='Related Model Name')
    ref_id = fields.Many2oneReference(
        string='Related Record ID',
        model_field='ref_model',
    )

    # Binary fields
    binary_field = fields.Binary(string='Binary File')
    image_field = fields.Image(string='Image')
    

class VbaJsonSampleLine(models.Model):
    _name = 'vba_json.sample.line'
    _description = 'VBA Json Sample Line for One2many'

    vba_json_sample_id = fields.Many2one('vba_json.sample', string='VBA Json Sample Reference', required=True, ondelete='cascade')
    name = fields.Char(string='Line Name', required=True)
    quantity = fields.Integer(string='Quantity', default=1)
