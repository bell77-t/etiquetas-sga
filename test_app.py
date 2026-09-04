import unittest
from pathlib import Path
from data_loader import DataManager, generate_tank_labels

class TestSGABusinessRules(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dm = DataManager()

    def test_summary_and_catalog(self):
        summary = self.dm.get_summary()
        self.assertGreater(summary["total_applications"], 0)
        self.assertGreater(summary["total_labels"], 0)
        self.assertGreater(summary["total_products_in_catalog"], 0)
        self.assertTrue(len(summary["available_dates"]) > 0)
        print(f"[TEST PASS] Catalogo cargado con {summary['total_applications']} apps y {summary['total_labels']} etiquetas.")

    def test_tank_breakdown_162600(self):
        # Caso 162,600 L con dosis 0.811
        mock_item = {
            "id": "app-test-1",
            "producto": "NITRATO DE CALCIO + ZN",
            "sector_bloque": "Sector 3 0",
            "reentrada": 0,
            "categoria": "N/A",
            "fecha": {"iso": "2024-03-04", "display": "04/03/2024"},
            "litros_total": 162600.0,
            "dosis": 0.811,
            "base_info": {
                "um": "KILO",
                "palabra_advertencia": "PELIGRO",
                "frase_h": "H302...",
                "frase_p": "P264...",
                "pictogramas": []
            }
        }
        labels = generate_tank_labels(mock_item)
        # 162600 // 1000 = 162 completos + 1 residuo de 600 = 163 etiquetas
        self.assertEqual(len(labels), 163)
        self.assertEqual(labels[0]["litros_tanque"], 1000.0)
        self.assertAlmostEqual(labels[0]["cantidad_dosificar"], 811.0, places=2)
        self.assertFalse(labels[0]["es_colita"])
        
        colita = labels[-1]
        self.assertTrue(colita["es_colita"])
        self.assertEqual(colita["litros_tanque"], 600.0)
        self.assertAlmostEqual(colita["cantidad_dosificar"], 600.0 * 0.811, places=2)
        print("[TEST PASS] Regla de 162,600 L desglose exacto (162 tanques 1,000L + 1 colita 600L).")

    def test_tank_breakdown_less_than_1000(self):
        # Caso 650 L con dosis 0.5
        mock_item = {
            "id": "app-test-2",
            "producto": "PRODUCTO PEQUEÑO",
            "sector_bloque": "Bloque 1",
            "reentrada": 12,
            "categoria": "III",
            "fecha": {"iso": "2024-03-05", "display": "05/03/2024"},
            "litros_total": 650.0,
            "dosis": 0.5,
            "base_info": {
                "um": "LITRO",
                "palabra_advertencia": "ATENCIÓN",
                "frase_h": "",
                "frase_p": "",
                "pictogramas": []
            }
        }
        labels = generate_tank_labels(mock_item)
        self.assertEqual(len(labels), 1)
        self.assertEqual(labels[0]["litros_tanque"], 650.0)
        self.assertEqual(labels[0]["cantidad_dosificar"], 325.0)
        self.assertFalse(labels[0]["es_colita"])
        print("[TEST PASS] Regla < 1,000 L genera etiqueta única con volumen exacto.")

    def test_pictogram_resolution(self):
        # Verificar que productos como ABAMECAL o ACEITE MINERAL tengan pictogramas válidos
        prod_abamecal = self.dm.base_catalog.get("ABAMECAL 1.8 EC")
        self.assertIsNotNone(prod_abamecal)
        pictos = prod_abamecal["pictogramas"]
        self.assertEqual(len(pictos), 4)
        has_imgs = [p["has_image"] for p in pictos]
        self.assertTrue(any(has_imgs))
        print(f"[TEST PASS] Pictogramas para ABAMECAL resueltos: {[p['filename'] for p in pictos if p['has_image']]}")

if __name__ == "__main__":
    unittest.main()
