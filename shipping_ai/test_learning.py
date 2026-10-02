import unittest
import gc
from contextlib import closing
import tempfile
from pathlib import Path
from database import init_db, get_connection
from learning import suggest_plan_from_history, validate_history_plan, physical_profile

class LearningTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.db = str(Path(self.temp.name)/'test.db')
        init_db(self.db)
        with closing(get_connection(self.db)) as c:
            for pid in range(1, 7):
                c.execute('INSERT INTO products(id,name,length_cm,width_cm,height_cm,weight) VALUES(?,?,?,?,?,?)', (pid,f'Item {pid}',1,2,3,.01))
            for bid in (1,2):
                c.execute('INSERT INTO boxes(id,name,length_cm,width_cm,height_cm,max_weight,is_active) VALUES(?,?,?,?,?,?,?)',(bid,f'Box {bid}',100,100,100,100,1))
            c.commit()
        self.boxes = [dict(id=b,length_cm=100,width_cm=100,height_cm=100,max_weight=100,is_active=1) for b in (1,2)]
    def tearDown(self):
        gc.collect()
        self.temp.cleanup()
    def item(self,pid,qty):
        return dict(product_id=pid,name=f'Item {pid}',length_cm=1,width_cm=2,height_cm=3,weight=.01,quantity=qty)
    def history(self,oid,items,box=1,assignments=True):
        import json
        with closing(get_connection(self.db)) as c:
            c.execute('INSERT INTO orders(id) VALUES(?)',(oid,))
            for pid,qty in items:
                c.execute('INSERT INTO order_items(order_id,product_id,quantity) VALUES(?,?,?)',(oid,pid,qty))
            c.execute('INSERT INTO shipment_history(order_id,box_id,quantity,assignments_json) VALUES(?,?,1,?)',(oid,box,json.dumps(dict((str(p),q) for p,q in items)) if assignments else None))
            c.commit()
    def suggest(self,items):
        self.audit=[]
        return suggest_plan_from_history(self.db,items,audit=self.audit,validator=lambda plan:validate_history_plan(plan,items,self.boxes))
    def test_equivalent_skus_and_manual_assignment(self):
        self.history(1,[(1,100),(2,100),(3,50)])
        result=self.suggest([self.item(4,150),self.item(5,100)])
        self.assertEqual(result['source'],'history_physical_exact')
        self.assertEqual(result['plan'][0]['assignments'],{'4':150,'5':100})
    def test_exact(self):
        self.history(1,[(1,200)])
        self.assertEqual(self.suggest([self.item(1,200)])['source'],'history_plan_exact')
    def test_less_and_more(self):
        self.history(1,[(1,200)])
        self.assertIsNotNone(self.suggest([self.item(2,180)]))
        self.assertIsNone(self.suggest([self.item(2,260)]))
        self.assertEqual(self.audit[0]['reason'],'quantity_above_reference')
    def test_conflicts(self):
        for oid in range(1,6): self.history(oid,[(1,200)],box=1 if oid<=3 else 2)
        result=self.suggest([self.item(1,200)])
        self.assertEqual(result['confidence'],60)
        self.assertEqual(result['conflict_count'],1)
    def test_inactive_and_capacity(self):
        self.history(1,[(1,200)])
        self.boxes[0]['is_active']=0
        self.assertIsNone(self.suggest([self.item(1,200)]))
        self.boxes[0]['is_active']=1
        self.boxes[0]['length_cm']=.5
        self.assertIsNone(self.suggest([self.item(1,200)]))
    def test_orientation_and_pack_family(self):
        a=self.item(1,10); b=self.item(2,10)
        b['length_cm'],b['height_cm']=3,1
        self.assertEqual(physical_profile(a),physical_profile(b))
        b['name']='SSD SATA 2.5'
        self.assertNotEqual(physical_profile(a),physical_profile(b))
    def test_additional_product_is_counted(self):
        self.history(1,[(1,200)])
        extra = self.item(6,1); extra['length_cm']=10
        result=self.suggest([self.item(1,200), extra])
        self.assertEqual(result['plan'][0]['assignments'], {'1':200, '6':1})
        self.boxes[0].update(length_cm=5,width_cm=5,height_cm=5)
        self.assertIsNone(self.suggest([self.item(1,200), extra]))

    def test_zero_assignments_from_interactive_ui(self):
        self.history(1,[(1,100),(2,100)])
        with closing(get_connection(self.db)) as c:
            c.execute('DELETE FROM shipment_history')
            c.execute("INSERT INTO shipment_history(order_id,box_id,quantity,assignments_json) VALUES(1,1,1,?)", ('{"1":100,"2":0}',))
            c.execute("INSERT INTO shipment_history(order_id,box_id,quantity,assignments_json) VALUES(1,2,1,?)", ('{"1":0,"2":100}',))
            c.commit()
        result=self.suggest([self.item(1,100),self.item(2,100)])
        self.assertEqual(len(result['plan']),2)

    def test_recommendation_uses_history_and_preserves_special_rule(self):
        import ast
        from packing import calculate_order_totals
        tree=ast.parse(Path('app_volum.py').read_text(encoding='utf-8'))
        function=next(n for n in tree.body if isinstance(n,ast.FunctionDef) and n.name=='build_recommendation')
        algorithm=dict(box=self.boxes[0],packages_required=1,reason='algorithm',confidence=70)
        scope=dict(DB_PATH=self.db, get_packing_rules=lambda db:{}, list_boxes=lambda db:self.boxes,
                   _resolve_mb_context=lambda *args:None, _resolve_ssdm2_context=lambda *args:None,
                   _build_effective_rules=lambda *args:({},None), calculate_order_totals=calculate_order_totals,
                   choose_best_box=lambda *args,**kwargs:algorithm,
                   _apply_mb_default_rule=lambda *args:(algorithm,False),
                   _apply_ssdm2_default_rule=lambda *args:(algorithm,False),
                   _with_weight_breakdown=lambda totals,*args:totals,
                   suggest_plan_from_history=suggest_plan_from_history,
                   validate_history_plan=validate_history_plan)
        import json
        scope['json']=json
        exec(compile(ast.Module(body=[function],type_ignores=[]),'app_volum.py','exec'),scope)
        self.history(1,[(1,100)],box=2)
        result=scope['build_recommendation']([self.item(1,100)])
        self.assertEqual(result['box']['id'],2)
        self.assertIn('assignments_json',result['history_plan'][0])
        scope['_apply_mb_default_rule']=lambda *args:(algorithm,True)
        result=scope['build_recommendation']([self.item(1,100)])
        self.assertEqual(result['box']['id'],1)
        self.assertEqual(result['learning_audit'][0]['reason'],'special_rule_incompatible')

    def test_operational_families(self):
        profiles=[]
        for name in ['SSD 2.5', 'SSD M.2', 'RAM notebook', 'RAM desktop', 'placa mae']:
            item=self.item(1,100); item['name']=name
            profiles.append(physical_profile(item))
        self.assertEqual(len(set(profiles)),5)

    def test_metrics(self):
        from learning_metrics import observe, feedback, report
        items=[dict(self.item(1,200),order_id=9)]
        rec=dict(box=self.boxes[0],packages_required=1,source='algorithm',evidence_count=0)
        observe(self.db,items,rec); observe(self.db,items,rec)
        feedback(self.db,9,items,[dict(box_id=2,quantity=1)])
        stats=report(self.db)
        self.assertEqual(stats['observed_orders'],1)
        self.assertEqual(stats['correction_percent'],100)

    def test_malformed_assignment(self):
        self.history(1,[(1,200)])
        with closing(get_connection(self.db)) as c:
            c.execute("UPDATE shipment_history SET assignments_json='{}'"); c.commit()
        self.assertIsNone(self.suggest([self.item(1,200)]))
        self.assertEqual(self.audit[0]['reason'],'incomplete_assignments')

if __name__=='__main__': unittest.main()
