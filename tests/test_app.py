import unittest
from pathlib import Path
from backend.data_loader import DataManager, generate_tank_labels
from backend import audit_logger

class TestSGABusinessRules(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._audit_db_path = audit_logger.DB_PATH
        cls._test_audit_db = Path("var") / "test_audit.db"
        cls._test_audit_db.unlink(missing_ok=True)
        audit_logger.DB_PATH = cls._test_audit_db
        cls.dm = DataManager()

    @classmethod
    def tearDownClass(cls):
        audit_logger.DB_PATH = cls._audit_db_path
        cls._test_audit_db.unlink(missing_ok=True)

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

    def test_programs_switching(self):
        # Verificar programas disponibles y cambio de programa activo
        programs = self.dm.programs
        self.assertIn("Data", programs)
        self.assertIn("ALZ", programs)
        self.assertIn("R-S-L", programs)
        self.assertEqual(programs["Data"]["count"], 90)
        self.assertEqual(programs["Data"]["total_labels"], 90)
        self.assertEqual(programs["ALZ"]["count"], 124)
        self.assertEqual(programs["R-S-L"]["count"], 134)

        # Switch to ALZ
        self.dm.set_active_program("ALZ")
        self.assertEqual(self.dm.current_program, "ALZ")
        self.assertEqual(len(self.dm.applications), 124)

        # Switch to R-S-L
        self.dm.set_active_program("R-S-L")
        self.assertEqual(self.dm.current_program, "R-S-L")
        self.assertEqual(len(self.dm.applications), 134)

        # Switch back to Data
        self.dm.set_active_program("Data")
        self.assertEqual(self.dm.current_program, "Data")
        self.assertEqual(len(self.dm.applications), 90)
        print("[TEST PASS] Soporte multi-programa verificado (Data: 90, ALZ: 124, R-S-L: 134).")

    def test_fastapi_endpoints(self):
        from fastapi.testclient import TestClient
        from app import app
        client = TestClient(app)

        res = client.get("/api/summary")
        self.assertEqual(res.status_code, 200)
        self.assertIn(res.json()["current_program"], ["Data", "MIPE", "MIRFE"])

        res_alz = client.post("/api/set-program", json={"program": "ALZ"})
        self.assertEqual(res_alz.status_code, 200)
        self.assertEqual(res_alz.json()["current_program"], "ALZ")
        self.assertEqual(len(res_alz.json()["applications"]), 124)

        # Reset to MIPE
        client.post("/api/set-program", json={"program": "MIPE"})
        print("[TEST PASS] Endpoints FastAPI para cambio de programa validados con TestClient.")

    def test_program_requests_do_not_mutate_shared_state(self):
        from fastapi.testclient import TestClient
        from app import app, data_manager
        client = TestClient(app)
        data_manager.set_active_program("Data")

        response = client.get("/api/applications?program=ALZ")
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["count"], 124)
        self.assertEqual(data_manager.current_program, "Data")

        invalid_copies = client.post("/api/export-pdf", json={"copies": 501})
        self.assertEqual(invalid_copies.status_code, 422)

    def test_ghs_inference(self):
        from backend.ghs_engine import infer_ghs_pictograms
        # H318 & H410
        res = infer_ghs_pictograms("H318 Provoca lesiones oculares graves. H410 Muy tóxico para organismos acuáticos.")
        codes = [p["code"] for p in res]
        self.assertIn("GHS05", codes)
        self.assertIn("GHS09", codes)
        print(f"[TEST PASS] Inferencia GHS verificada: H318->GHS05, H410->GHS09: {codes}")

    def test_audit_logger(self):
        log_id = audit_logger.log_event(
            operario="Inspector Test",
            accion="IMPRESION_TEST",
            programa="ALZ",
            producto="CAPTAN 480 SC",
            copias=3
        )
        self.assertIsNotNone(log_id)
        logs = audit_logger.get_logs(limit=5)
        self.assertGreaterEqual(logs["total"], 1)
        self.assertTrue(any(l["operario"] == "Inspector Test" for l in logs["items"]))
        print("[TEST PASS] Registro y trazabilidad en SQLite verificado.")

    def test_pdf_generation(self):
        from backend import pdf_generator
        labels = self.dm.applications[0]["etiquetas"][:2]
        picto_dir = self.dm.picto_dir
        buf = pdf_generator.generate_pdf(labels, picto_dir, layout="letter")
        self.assertGreater(len(buf.getvalue()), 1000)
        buf_th = pdf_generator.generate_pdf(labels, picto_dir, layout="thermal")
        self.assertGreater(len(buf_th.getvalue()), 1000)
        print("[TEST PASS] Generación de PDF vectorial (Carta y Rollo Térmico) verificada.")

    def test_advanced_api_endpoints(self):
        from fastapi.testclient import TestClient
        from app import app
        client = TestClient(app)

        # 1. Network Info
        net = client.get("/api/network-info")
        self.assertEqual(net.status_code, 200)
        self.assertIn("local_ip", net.json())

        # 2. QR Code
        qr = client.get("/api/qr/1ACABA01")
        self.assertEqual(qr.status_code, 200)
        self.assertEqual(qr.headers["content-type"], "image/png")

        # 3. Ficha FDS HTML
        ficha = client.get("/ficha/1ACABA01")
        self.assertEqual(ficha.status_code, 200)
        self.assertIn("FICHA BREVE SGA", ficha.text.upper())
        self.assertIn("IMPRIMIR FICHA", ficha.text.upper())
        self.assertNotIn("CISPROQUIM", ficha.text)

        # 4. Audit API
        audit_res = client.get("/api/audit/logs")
        self.assertEqual(audit_res.status_code, 200)

        # 5. PDF Export API
        apps = client.get("/api/applications").json()["items"]
        app_id = apps[0]["id"]
        pdf_res = client.post("/api/export-pdf", json={"application_ids": [app_id], "layout": "thermal"})
        self.assertEqual(pdf_res.status_code, 200)
        self.assertEqual(pdf_res.headers["content-type"], "application/pdf")
        print("[TEST PASS] Endpoints avanzados validados (Red Local, QR PNG, Ficha HTML, Auditoría, PDF Export).")

if __name__ == "__main__":
    unittest.main()
