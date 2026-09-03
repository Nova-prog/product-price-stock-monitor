import unittest

from monitor import (
    Product,
    build_comparison_rows,
    parse_askul,
    parse_kanto,
    parse_monotaro,
    parse_wako,
)


CHECKED_AT = "2026-09-03T09:00:00+09:00"


class MonitorParserTests(unittest.TestCase):
    def test_monotaro_price_and_unit_price(self):
        product = Product(
            enabled=True,
            group="gloves",
            category="lab_consumable",
            site="monotaro",
            name="Nitrile gloves",
            url="https://example.test/monotaro",
            unit="1 box",
            quantity="100",
            quantity_unit="枚",
        )
        result = parse_monotaro(product, "<title>Test</title>販売価格 ￥1,280 在庫あり 明日出荷", CHECKED_AT)
        self.assertEqual(result.price, 1280)
        self.assertEqual(result.unit_price, 12.8)
        self.assertEqual(result.price_type, "販売価格")

    def test_askul_price_after_label(self):
        product = Product(True, "kimwipes", "lab_consumable", "askul", "Kimwipes", "https://example.test")
        result = parse_askul(product, "税込価格 742円 お届け 明日", CHECKED_AT)
        self.assertEqual(result.price, 742)
        self.assertEqual(result.delivery_status, "お届け 明日")

    def test_kanto_uses_list_price(self):
        product = Product(True, "reagent", "reagent", "kanto", "Acetone", "https://example.test")
        result = parse_kanto(product, "定価 1,900円 在庫 あり", CHECKED_AT)
        self.assertEqual(result.price, 1900)
        self.assertEqual(result.price_type, "定価")

    def test_wako_uses_suggested_price(self):
        product = Product(True, "reagent", "reagent", "wako", "Acetone", "https://example.test")
        result = parse_wako(product, "希望納入価格 2,500円 東日本 在庫あり", CHECKED_AT)
        self.assertEqual(result.price, 2500)
        self.assertEqual(result.price_type, "希望納入価格")

    def test_comparison_prefers_lower_unit_price(self):
        rows = [
            {
                "checked_at": CHECKED_AT,
                "group": "gloves",
                "category": "lab_consumable",
                "site": "monotaro",
                "name": "A",
                "price": 1200,
                "unit_price": 12,
                "currency": "JPY",
                "tax_status": "税別",
                "price_type": "販売価格",
                "result_state": "ok",
            },
            {
                "checked_at": CHECKED_AT,
                "group": "gloves",
                "category": "lab_consumable",
                "site": "askul",
                "name": "B",
                "price": 1100,
                "unit_price": 11,
                "currency": "JPY",
                "tax_status": "税別",
                "price_type": "販売価格",
                "result_state": "ok",
            },
        ]
        comparison = build_comparison_rows(rows, CHECKED_AT)
        self.assertEqual(comparison[0]["best_site"], "askul")


if __name__ == "__main__":
    unittest.main()
