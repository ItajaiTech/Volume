import ast
import unittest
from pathlib import Path
from functools import wraps
from flask import Flask, request, session, redirect, url_for
from learning import validate_history_plan
from packing import calculate_order_totals


class VolumetryValidationTests(unittest.TestCase):
    def setUp(self):
        self.items = [dict(product_id=30, name="MEM DESK DDR4 16GB", quantity=380,
                           length_cm=15.9, width_cm=4.6, height_cm=.9, weight=.028)]
        self.boxes = [dict(id=6, name="M2", length_cm=30, width_cm=30, height_cm=30,
                          max_weight=.25, max_capacity_weight=30, is_active=1),
                      dict(id=10, name="MB", length_cm=44.7, width_cm=20.5, height_cm=49.5,
                           max_weight=.93, max_capacity_weight=30, is_active=1)]
        app = Flask(__name__)
        app.secret_key = "test"
        app.add_url_rule('/login', 'admin_login', lambda: 'login')
        scope = dict(app=app, request=request, session=session, redirect=redirect,
                     url_for=url_for, wraps=wraps, DB_PATH='unused',
                     ORDER_NOT_FOUND_MSG='Pedido nao encontrado.',
                     get_order=lambda db, oid: {'id':149} if oid == 149 else None,
                     get_order_items=lambda *args:self.items,
                     list_boxes=lambda *args:self.boxes, get_packing_rules=lambda *args:{},
                     validate_history_plan=validate_history_plan, calculate_order_totals=calculate_order_totals)
        names = {'admin_required', '_volumetry_error_message', 'order_volumetry_validate'}
        tree = ast.parse(Path('app_volum.py').read_text(encoding='utf-8'))
        nodes = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name in names]
        exec(compile(ast.Module(body=nodes, type_ignores=[]), 'app_volum.py', 'exec'), scope)
        self.client = app.test_client()
        with self.client.session_transaction() as sess:
            sess['is_admin'] = True

    def validate(self, packages):
        return self.client.post('/order/149/volumetry/validate', json=packages)

    def package(self, box, qty):
        return dict(box_id=box, assignments={'30':qty})

    def test_m2_rejected_despite_raw_volume(self):
        self.assertLess(380*15.9*4.6*.9, 30**3)
        result = self.validate([self.package(6,380)]).get_json()
        self.assertFalse(result['valid'])
        self.assertFalse(result['packages'][0]['valid'])
        self.assertIn('packs', result['message'])
        self.assertNotIn('physical_or_pack', result['message'])

    def test_mb_accepted(self):
        result = self.validate([self.package(10,380)]).get_json()
        self.assertTrue(result['valid'])
        self.assertTrue(result['packages'][0]['valid'])

    def test_split_m2(self):
        result = self.validate([self.package(6,180),self.package(6,180),self.package(6,20)]).get_json()
        self.assertTrue(result['valid'])
        self.assertTrue(all(p['valid'] for p in result['packages']))

    def test_partial_plan_can_have_valid_package(self):
        result = self.validate([self.package(6,100)]).get_json()
        self.assertFalse(result['valid'])
        self.assertTrue(result['packages'][0]['valid'])

    def test_weight_capacity(self):
        self.boxes[1]['max_capacity_weight'] = 10
        self.assertFalse(self.validate([self.package(10,380)]).get_json()['valid'])

    def test_bad_input_and_auth(self):
        for payload in ({}, [None]):
            self.assertEqual(self.validate(payload).status_code,400)
        for payload in ([self.package(None,380)], [self.package(10,-1)], [self.package(10,380.5)]):
            self.assertFalse(self.validate(payload).get_json()['valid'])
        self.assertEqual(self.client.post('/order/999/volumetry/validate', json=[]).status_code,404)
        with self.client.session_transaction() as sess:
            sess.clear()
        self.assertEqual(self.validate([self.package(10,380)]).status_code,302)


if __name__ == '__main__':
    unittest.main()
