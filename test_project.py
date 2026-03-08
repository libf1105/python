"""
单元测试 - 覆盖核心业务逻辑
"""
import unittest
import os
import tempfile
import pandas as pd
from unittest.mock import patch


# ─────────────────────────────────────────────
# utils 模块测试
# ─────────────────────────────────────────────
from utils import (
    analyze_goods_type,
    analyze_import_export_type,
    analyze_insurance_type,
    extract_bill_number,
    generate_policy_filename,
)


class TestAnalyzeGoodsType(unittest.TestCase):
    def test_dangerous_battery(self):
        self.assertEqual(analyze_goods_type("锂电池 battery"), "危险品")

    def test_dangerous_chemical(self):
        self.assertEqual(analyze_goods_type("化学品 chemical"), "危险品")

    def test_fragile_glass(self):
        self.assertEqual(analyze_goods_type("玻璃制品"), "易碎品")

    def test_fragile_ceramic(self):
        self.assertEqual(analyze_goods_type("陶瓷工艺品"), "易碎品")

    def test_normal_goods(self):
        self.assertEqual(analyze_goods_type("普通货物"), "普货")

    def test_empty_input(self):
        self.assertEqual(analyze_goods_type(""), "普货")

    def test_none_input(self):
        self.assertEqual(analyze_goods_type(None), "普货")

    def test_dangerous_takes_priority_over_fragile(self):
        # 同时含危险品和易碎品关键词，危险品优先
        self.assertEqual(analyze_goods_type("玻璃电池"), "危险品")


class TestAnalyzeImportExportType(unittest.TestCase):
    def test_china_export(self):
        self.assertEqual(analyze_import_export_type("中国上海"), "出口")

    def test_china_english_export(self):
        self.assertEqual(analyze_import_export_type("China Shanghai"), "出口")

    def test_cn_export(self):
        self.assertEqual(analyze_import_export_type("CN"), "出口")

    def test_foreign_import(self):
        self.assertEqual(analyze_import_export_type("美国"), "进口")

    def test_empty_defaults_export(self):
        self.assertEqual(analyze_import_export_type(""), "出口")

    def test_none_defaults_export(self):
        self.assertEqual(analyze_import_export_type(None), "出口")


class TestAnalyzeInsuranceType(unittest.TestCase):
    def test_express_keyword_shunfeng(self):
        self.assertEqual(analyze_insurance_type("顺丰快递", ""), "快递")

    def test_express_keyword_in_subject(self):
        self.assertEqual(analyze_insurance_type("", "快递投保申请"), "快递")

    def test_express_english(self):
        self.assertEqual(analyze_insurance_type("express delivery", ""), "快递")

    def test_non_express(self):
        self.assertEqual(analyze_insurance_type("海运货物", "普通货运"), "非快递")

    def test_empty_input(self):
        self.assertEqual(analyze_insurance_type("", ""), "非快递")


class TestExtractBillNumber(unittest.TestCase):
    def test_chinese_bill_colon(self):
        self.assertEqual(extract_bill_number("提单号：ABC12345678"), "ABC12345678")

    def test_chinese_bill_no_colon(self):
        self.assertEqual(extract_bill_number("提单号 ABC12345678"), "ABC12345678")

    def test_bl_format(self):
        self.assertEqual(extract_bill_number("B/L: ABC12345678"), "ABC12345678")

    def test_uppercase_pattern(self):
        self.assertEqual(extract_bill_number("单号 ABC12345678"), "ABC12345678")

    def test_pure_digits(self):
        result = extract_bill_number("单号 123456789012")
        self.assertEqual(result, "123456789012")

    def test_no_bill_number(self):
        self.assertEqual(extract_bill_number("这封邮件没有提单号"), "")


class TestGeneratePolicyFilename(unittest.TestCase):
    def test_default_pdf(self):
        name = generate_policy_filename("POL123")
        self.assertTrue(name.startswith("POL123_"))
        self.assertTrue(name.endswith(".pdf"))

    def test_custom_extension(self):
        name = generate_policy_filename("POL123", "xlsx")
        self.assertTrue(name.endswith(".xlsx"))


# ─────────────────────────────────────────────
# database 模块测试（使用临时文件隔离）
# ─────────────────────────────────────────────
class TestLocalSheetManager(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.excel_path = os.path.join(self.tmp_dir, "投保记录.xlsx")
        # 将 FILE_CONFIG 指向临时路径
        self.patcher = patch(
            "database.FILE_CONFIG",
            {"excel_path": self.excel_path}
        )
        self.patcher.start()
        from database import LocalSheetManager
        self.manager = LocalSheetManager()

    def tearDown(self):
        self.patcher.stop()

    def test_excel_created_on_init(self):
        self.assertTrue(os.path.exists(self.manager.excel_path))

    def test_add_and_get_record(self):
        self.manager.add_record({"邮件标题": "测试邮件", "发件人": "a@b.com"})
        records = self.manager.get_all_records()
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["邮件标题"], "测试邮件")

    def test_add_record_auto_status(self):
        self.manager.add_record({"邮件标题": "测试"})
        records = self.manager.get_all_records()
        self.assertEqual(records[0]["投保状态"], "未投保")
        self.assertEqual(records[0]["邮件回复状态"], "未回复")

    def test_record_exists_true(self):
        self.manager.add_record({
            "邮件接收时间": "2024-01-01 12:00:00",
            "发件人": "a@b.com",
            "邮件标题": "测试"
        })
        self.assertTrue(
            self.manager.record_exists("2024-01-01 12:00:00", "a@b.com", "测试")
        )

    def test_record_exists_false(self):
        self.assertFalse(
            self.manager.record_exists("2024-01-01 12:00:00", "x@y.com", "不存在")
        )

    def test_update_record(self):
        self.manager.add_record({"邮件标题": "更新测试", "序号": 1})
        result = self.manager.update_record("邮件标题", "更新测试", {"投保状态": "已投保"})
        self.assertTrue(result)
        records = self.manager.get_all_records()
        self.assertEqual(records[0]["投保状态"], "已投保")

    def test_update_record_not_found(self):
        result = self.manager.update_record("邮件标题", "不存在的标题", {"投保状态": "已投保"})
        self.assertFalse(result)

    def test_get_uninsured_records(self):
        self.manager.add_record({"邮件标题": "未投保邮件"})
        self.manager.add_record({"邮件标题": "已投保邮件"})
        # 手动将第二条改为已投保
        self.manager.update_record("邮件标题", "已投保邮件", {"投保状态": "已投保"})
        uninsured = self.manager.get_uninsured_records()
        self.assertEqual(len(uninsured), 1)
        self.assertEqual(uninsured[0]["邮件标题"], "未投保邮件")

    def test_get_all_records_empty(self):
        records = self.manager.get_all_records()
        self.assertEqual(records, [])

    def test_sequential_row_numbers(self):
        self.manager.add_record({"邮件标题": "第一封"})
        self.manager.add_record({"邮件标题": "第二封"})
        records = self.manager.get_all_records()
        self.assertEqual(records[0]["序号"], 1)
        self.assertEqual(records[1]["序号"], 2)


# ─────────────────────────────────────────────
# update_insurance_record 模块测试
# ─────────────────────────────────────────────
class TestUpdateInsuranceStatus(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.excel_path = os.path.join(self.tmp_dir, "投保记录.xlsx")
        # 创建测试用 Excel
        df = pd.DataFrame([
            {"序号": 1, "邮件标题": "邮件A", "投保状态": "未投保", "保单号": ""},
            {"序号": 2, "邮件标题": "邮件B", "投保状态": "未投保", "保单号": ""},
        ])
        df.to_excel(self.excel_path, index=False, engine="openpyxl")
        self.patcher = patch(
            "update_insurance_record.FILE_CONFIG",
            {"excel_path": self.excel_path}
        )
        self.patcher.start()
        from update_insurance_record import update_insurance_status
        self.update_fn = update_insurance_status

    def tearDown(self):
        self.patcher.stop()

    def test_update_success(self):
        result = self.update_fn(1, "POL-2024-001")
        self.assertTrue(result)
        df = pd.read_excel(self.excel_path, engine="openpyxl")
        row = df[df["序号"] == 1].iloc[0]
        self.assertEqual(row["保单号"], "POL-2024-001")
        self.assertEqual(row["投保状态"], "已投保")

    def test_update_does_not_affect_other_rows(self):
        self.update_fn(1, "POL-2024-001")
        df = pd.read_excel(self.excel_path, engine="openpyxl")
        row2 = df[df["序号"] == 2].iloc[0]
        self.assertEqual(row2["投保状态"], "未投保")

    def test_row_not_found(self):
        result = self.update_fn(99, "POL-XXXX")
        self.assertFalse(result)

    def test_file_not_exist(self):
        result = self.update_fn.__wrapped__(999, "POL") if hasattr(self.update_fn, "__wrapped__") else None
        # 直接测试文件不存在场景
        with patch("update_insurance_record.FILE_CONFIG", {"excel_path": "不存在的路径.xlsx"}):
            from update_insurance_record import update_insurance_status as fn
            self.assertFalse(fn(1, "POL"))


if __name__ == "__main__":
    unittest.main(verbosity=2)
