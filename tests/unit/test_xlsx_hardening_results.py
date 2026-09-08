"""汇总协议测试；这里的行是统计单测夹具，不是Benchmark题目。"""
import sys,tempfile,unittest
from pathlib import Path
from unittest.mock import patch
sys.path.insert(0,str(Path(__file__).resolve().parents[2]/'tools'))
from xlsx_hardening_results import summarize
import json

class ResultsTest(unittest.TestCase):
    def test_partial_does_not_publish_difficulty(self):
        with tempfile.TemporaryDirectory() as d:
            p=Path(d);s=summarize(p,p,[{'task':'a','reward':-.2}],[])
            self.assertEqual(s['valid_count'],1);self.assertEqual(s['mean_reward'],-.2)
            self.assertFalse((p/'difficulty.json').exists())

    def test_complete_quantiles_keep_negative_and_ties(self):
        rows=[{'task':f't{i:02}','reward':.5 if i<48 else -.2,'origin':'fixture'} for i in range(60)]
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)
            with patch.object(Path,'read_text',return_value='[metadata]\nestimated_minutes=10\ntask_family="pivot_analysis"'):
                s=summarize(p,p,rows[:53],rows[53:])
            report=json.loads((p/'difficulty.json').read_text())
            self.assertEqual([x['count'] for x in report['tiers']],[12,18,18,12])
            self.assertAlmostEqual(report['tiers'][-1]['mean_reward'],-.2)
            self.assertEqual(report['boundary_ties'],[12,30])
            self.assertEqual(s['negative_count'],12)
            self.assertEqual(report['tasks'][0]['task'],'t00')

if __name__=='__main__':unittest.main()
