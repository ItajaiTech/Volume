import tempfile
import unittest
from pathlib import Path
from jinja2 import Environment, DictLoader
from database import DEFAULT_PACKING_RULES, EXTRA_PACKING_FIELDS, init_db, update_packing_rules, get_packing_rules
from packing import _build_preview_groups_for_item, calculate_order_totals, build_packing_3d_preview
from learning import validate_history_plan, physical_profile

class OverlapTests(unittest.TestCase):
    def setUp(self):
        self.rules = dict(DEFAULT_PACKING_RULES, ram_desk_bundle_length_cm=0,
                          ram_desk_bundle_width_cm=0, ram_desk_bundle_height_cm=7.5,
                          ram_overlap_enabled=1, ram_nb_bundle_height_cm=7.5)
        self.item = dict(product_id=30, sku='PRD00017', name='MEM DESK DDR4 16GB',
                         length_cm=15.9, width_cm=4.6, height_cm=.9, quantity=380, weight=.028)
        self.box = dict(id=6, name='M2', length_cm=30, width_cm=30, height_cm=30,
                        max_weight=.25, max_capacity_weight=30, is_active=1)

    def reason(self, item=None):
        item = item or self.item
        return validate_history_plan([dict(box_id=6, quantity=1, assignments={'30':item['quantity']})],
                                     [item], [self.box], self.rules)

    def test_order149_and_preview(self):
        self.assertIsNone(self.reason())
        groups, _ = _build_preview_groups_for_item(self.item, self.rules)
        self.assertEqual([(g['blocks'],g['dims']) for g in groups],[(19,[29.6,4.6,7.5])])
        self.assertAlmostEqual(calculate_order_totals([self.item],self.rules)['total_volume_cm3'],19402.8)
        preview=build_packing_3d_preview([self.item],self.box,self.rules)
        self.assertEqual(preview['placed_blocks'],19)
        self.assertEqual(preview['overflow_blocks'],0)
        self.assertFalse(preview['unpack_applied'])
        self.assertEqual(self.item['length_cm'],15.9)

    def test_odd_pack_and_loose_units(self):
        item=dict(self.item,quantity=395)
        groups,_ = _build_preview_groups_for_item(item,self.rules)
        self.assertEqual([(g['blocks'],g['dims']) for g in groups],
                         [(19,[29.6,4.6,7.5]),(1,[15.9,4.6,7.5]),(5,[15.9,4.6,.9])])

    def test_scope_disable_and_edit(self):
        self.assertIsNone(self.reason(dict(self.item,sku='OTHER')))
        self.assertEqual(physical_profile(self.item,self.rules),physical_profile(dict(self.item,sku='OTHER'),self.rules))
        nb=dict(self.item,name='MEM NB DDR4',length_cm=9.5)
        groups,_=_build_preview_groups_for_item(nb,self.rules)
        self.assertAlmostEqual(groups[0]['dims'][0],16.8)
        self.assertEqual(groups[0]['dims'][2],7.5)
        self.rules['unpack_packs_single_volume']=0
        self.rules['ram_overlap_enabled']=0
        self.assertIsNotNone(self.reason())
        self.rules['ram_overlap_enabled']=1
        self.rules['ram_blister_tab_cm']=0
        self.rules['ram_pair_clearance_cm']=0
        self.assertIsNotNone(self.reason())

    def test_settings_roundtrip_and_panel_coverage(self):
        with tempfile.TemporaryDirectory() as tmp:
            db=str(Path(tmp)/'test.db');init_db(db)
            update_packing_rules(db,self.rules)
            loaded=get_packing_rules(db)
            self.assertEqual(loaded,self.rules)
            loaded['ram_blister_tab_cm']=1.1
            update_packing_rules(db,loaded)
            self.assertEqual(get_packing_rules(db)['ram_blister_tab_cm'],1.1)
        env=Environment(loader=DictLoader({'dashboard.html':Path('templates/dashboard.html').read_text(encoding='utf-8'),
                                          'base.html':'{% block content %}{% endblock %}'}))
        html=env.get_template('dashboard.html').render(stats=dict(most_used_box=None,avg_volume=0,top_products=[]),
            packing_rules=self.rules,extra_packing_fields=EXTRA_PACKING_FIELDS,url_for=lambda *a,**k:'#')
        for key in DEFAULT_PACKING_RULES:
            self.assertIn('name="'+key+'"',html,key)

if __name__ == '__main__': unittest.main()
