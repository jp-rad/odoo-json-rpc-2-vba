# from odoo import http


# class VbaJson(http.Controller):
#     @http.route('/vba_json/vba_json', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/vba_json/vba_json/objects', auth='public')
#     def list(self, **kw):
#         return http.request.render('vba_json.listing', {
#             'root': '/vba_json/vba_json',
#             'objects': http.request.env['vba_json.vba_json'].search([]),
#         })

#     @http.route('/vba_json/vba_json/objects/<model("vba_json.vba_json"):obj>', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('vba_json.object', {
#             'object': obj
#         })

