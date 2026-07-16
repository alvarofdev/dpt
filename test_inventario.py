"""
Tests de integración para el sistema de inventario de rollos.

Ejecutar con:
    pytest test_inventario.py -v

Requiere estar en el mismo directorio que inventario.py y checksum_rollo.py.
"""
import json
import os
import pytest
from datetime import datetime

from inventario import (
    CREDENTIALS,
    INVENTARIO_PREFIX,
    SHEET_REGISTROS,
    SHEET_UBICACIONES,
    ExcelManager,
    RegistroRollo,
    Ubicacion,
)
from checksum_rollo import calcular_checksum, validar_lote


# ---------------------------------------------------------------------------
# Fixtures compartidas
# ---------------------------------------------------------------------------

@pytest.fixture
def em(tmp_path):
    """ExcelManager con archivo temporal. Cada test recibe instancia limpia."""
    return ExcelManager(str(tmp_path / "test.xlsx"))


@pytest.fixture
def em_con_ubicacion(em):
    """ExcelManager con una ubicación ya creada."""
    em.add_ubicacion(Ubicacion("DEP1", "A1", "47", "02"))
    return em


# ---------------------------------------------------------------------------
# Modelo Ubicacion
# ---------------------------------------------------------------------------

class TestUbicacion:

    def test_id_concatenacion_directa(self):
        u = Ubicacion("DEP1", "A1", "47", "02")
        assert u.id == "DEP1A14702"

    def test_id_normaliza_a_mayusculas(self):
        u = Ubicacion("dep1", "a1", "47", "02")
        assert u.id == "DEP1A14702"

    def test_id_longitud_estandar_es_10(self):
        u = Ubicacion("DEP1", "A1", "47", "02")
        assert len(u.id) == 10

    def test_to_row_estructura(self):
        u = Ubicacion("DEP1", "A1", "47", "02")
        assert u.to_row() == ["DEP1A14702", "DEP1", "A1", "47", "02"]

    def test_to_row_longitud(self):
        u = Ubicacion("DEP1", "A1", "47", "02")
        assert len(u.to_row()) == 5

    def test_deposito_pasillo_columna_nivel_en_row(self):
        u = Ubicacion("DEPO", "B2", "03", "AA")
        row = u.to_row()
        assert row[1] == "DEPO"
        assert row[2] == "B2"
        assert row[3] == "03"
        assert row[4] == "AA"


# ---------------------------------------------------------------------------
# Modelo RegistroRollo
# ---------------------------------------------------------------------------

class TestRegistroRollo:

    def test_to_row_estructura(self):
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702", "2026-07-13 10:00:00")
        assert r.to_row() == ["user1", "FJLH78693", "DEP1A14702", "2026-07-13 10:00:00"]

    def test_fecha_default_no_vacia(self):
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702")
        assert r.fecha
        assert len(r.fecha) > 0

    def test_fecha_default_formato_fecha(self):
        before = datetime.now().strftime("%Y-%m-%d")
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702")
        assert before in r.fecha


# ---------------------------------------------------------------------------
# ExcelManager — inicialización
# ---------------------------------------------------------------------------

class TestExcelManagerInit:

    def test_crea_archivo_si_no_existe(self, tmp_path):
        filepath = str(tmp_path / "nuevo.xlsx")
        assert not os.path.exists(filepath)
        ExcelManager(filepath)
        assert os.path.exists(filepath)

    def test_crea_hoja_ubicaciones(self, tmp_path):
        em = ExcelManager(str(tmp_path / "nuevo.xlsx"))
        assert SHEET_UBICACIONES in em.wb.sheetnames

    def test_crea_hoja_registros(self, tmp_path):
        em = ExcelManager(str(tmp_path / "nuevo.xlsx"))
        assert SHEET_REGISTROS in em.wb.sheetnames

    def test_abre_archivo_existente_sin_duplicar_hojas(self, tmp_path):
        filepath = str(tmp_path / "nuevo.xlsx")
        ExcelManager(filepath)
        em2 = ExcelManager(filepath)
        assert em2.wb.sheetnames.count(SHEET_UBICACIONES) == 1
        assert em2.wb.sheetnames.count(SHEET_REGISTROS) == 1

    def test_get_ubicaciones_vacio_en_inicio(self, em):
        assert em.get_ubicaciones() == []

    def test_get_registros_vacio_en_inicio(self, em):
        assert em.get_registros() == []


# ---------------------------------------------------------------------------
# ExcelManager — ubicaciones
# ---------------------------------------------------------------------------

class TestAddUbicacion:

    def test_agrega_ubicacion_nueva(self, em):
        u = Ubicacion("DEP1", "A1", "47", "02")
        ok, msg = em.add_ubicacion(u)
        assert ok is True
        assert "DEP1A14702" in msg

    def test_rechaza_ubicacion_duplicada(self, em):
        u = Ubicacion("DEP1", "A1", "47", "02")
        em.add_ubicacion(u)
        ok, msg = em.add_ubicacion(u)
        assert ok is False
        assert "ya existe" in msg

    def test_ubicacion_recuperable_tras_guardar(self, em):
        em.add_ubicacion(Ubicacion("DEP1", "A1", "47", "02"))
        assert any(u.id == "DEP1A14702" for u in em.get_ubicaciones())

    def test_ubicacion_exists_true(self, em):
        em.add_ubicacion(Ubicacion("DEP1", "A1", "47", "02"))
        assert em.ubicacion_exists("DEP1A14702") is True

    def test_ubicacion_exists_false(self, em):
        assert em.ubicacion_exists("XXXXXXXXXX") is False

    def test_ubicacion_exists_case_insensitive(self, em):
        em.add_ubicacion(Ubicacion("DEP1", "A1", "47", "02"))
        assert em.ubicacion_exists("dep1a14702") is True

    def test_multiples_ubicaciones(self, em):
        em.add_ubicacion(Ubicacion("DEP1", "A1", "47", "02"))
        em.add_ubicacion(Ubicacion("DEP2", "B2", "10", "01"))
        ubicaciones = em.get_ubicaciones()
        ids = {u.id for u in ubicaciones}
        assert "DEP1A14702" in ids
        assert "DEP2B21001" in ids

    def test_get_ubicacion_by_id_encontrada(self, em):
        em.add_ubicacion(Ubicacion("DEP1", "A1", "47", "02"))
        u = em.get_ubicacion_by_id("DEP1A14702")
        assert u is not None
        assert u.id == "DEP1A14702"

    def test_get_ubicacion_by_id_no_encontrada(self, em):
        assert em.get_ubicacion_by_id("XXXXXXXXXX") is None


# ---------------------------------------------------------------------------
# ExcelManager — registros de rollos
# ---------------------------------------------------------------------------

class TestAddRegistro:

    @pytest.fixture(autouse=True)
    def setup(self, em_con_ubicacion):
        self.em = em_con_ubicacion

    def test_agrega_registro_nuevo(self):
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702")
        ok, msg = self.em.add_registro(r)
        assert ok is True
        assert "FJLH78693" in msg

    def test_rechaza_lote_duplicado(self):
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702")
        self.em.add_registro(r)
        ok, msg = self.em.add_registro(r)
        assert ok is False
        assert "ya fue registrado" in msg

    def test_lote_exists_true_tras_registro(self):
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702")
        self.em.add_registro(r)
        assert self.em.lote_exists("FJLH78693") is True

    def test_lote_exists_false_sin_registro(self):
        assert self.em.lote_exists("FJLH78693") is False

    def test_registro_recuperable_tras_guardar(self):
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702")
        self.em.add_registro(r)
        registros = self.em.get_registros()
        assert any(x.id_lote == "FJLH78693" for x in registros)

    def test_registros_de_diferentes_lotes(self):
        lotes = ["FJLH78693", "ABC123476", "XYZ999988"]
        for lote in lotes:
            r = RegistroRollo("user1", lote, "DEP1A14702")
            ok, _ = self.em.add_registro(r)
            assert ok is True
        assert len(self.em.get_registros()) == 3

    def test_rechazo_persiste_en_disco(self):
        r = RegistroRollo("user1", "FJLH78693", "DEP1A14702")
        self.em.add_registro(r)
        # Crear nueva instancia del manager apuntando al mismo archivo
        em2 = ExcelManager(self.em.filepath)
        ok, _ = em2.add_registro(r)
        assert ok is False


# ---------------------------------------------------------------------------
# ExcelManager — inventario mensual
# ---------------------------------------------------------------------------

class TestInventario:

    def test_crear_inventario_nuevo(self, em):
        ok, msg = em.crear_inventario("INV_2026-07")
        assert ok is True
        assert "INV_2026-07" in em.get_inventarios()

    def test_crear_inventario_con_prefix_automatico(self, em):
        # Si se pasa sin prefix, crear_inventario no lo agrega — eso lo hace la pantalla TUI
        ok, _ = em.crear_inventario("INV_2026-07")
        assert ok is True

    def test_crear_inventario_duplicado(self, em):
        em.crear_inventario("INV_2026-07")
        ok, msg = em.crear_inventario("INV_2026-07")
        assert ok is False
        assert "ya existe" in msg

    def test_get_inventarios_retorna_solo_prefixados(self, em):
        em.crear_inventario("INV_2026-07")
        em.crear_inventario("INV_2026-08")
        inventarios = em.get_inventarios()
        assert all(s.startswith(INVENTARIO_PREFIX) for s in inventarios)
        assert len(inventarios) == 2

    def test_add_a_inventario_persiste(self, em_con_ubicacion):
        em = em_con_ubicacion
        em.crear_inventario("INV_2026-07")
        em.add_a_inventario("INV_2026-07", "DEP1A14702", "FJLH78693", "user1")
        assert em.lote_en_inventario("INV_2026-07", "FJLH78693") is True

    def test_lote_no_en_inventario(self, em):
        em.crear_inventario("INV_2026-07")
        assert em.lote_en_inventario("INV_2026-07", "FJLH78693") is False

    def test_lote_en_inventario_hoja_inexistente(self, em):
        assert em.lote_en_inventario("INV_NOEXISTE", "FJLH78693") is False

    def test_add_a_inventario_sin_ubicacion_conocida(self, em):
        """Si la ubicación no está en el sistema, igual registra con datos vacíos."""
        em.crear_inventario("INV_2026-07")
        em.add_a_inventario("INV_2026-07", "XXXXXXXXXXX", "FJLH78693", "user1")
        assert em.lote_en_inventario("INV_2026-07", "FJLH78693") is True


# ---------------------------------------------------------------------------
# ExcelManager — import_from_android
# ---------------------------------------------------------------------------

class TestImportFromAndroid:

    def _make_json(self, tmp_path, registros=None, inventarios=None):
        data = {
            "registros": registros or [],
            "inventarios": inventarios or {},
        }
        path = str(tmp_path / "export.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f)
        return path

    def test_importa_registros_nuevos(self, em_con_ubicacion, tmp_path):
        json_path = self._make_json(tmp_path, registros=[
            {"id_usuario": "123", "id_lote": "FJLH78693",
             "id_ubicacion": "DEP1A14702", "fecha": "2026-07-13 10:00:00"},
        ])
        result = em_con_ubicacion.import_from_android(json_path)
        assert result["reg_nuevos"] == 1
        assert result["reg_dup"] == 0

    def test_detecta_registros_duplicados(self, em_con_ubicacion, tmp_path):
        em_con_ubicacion.add_registro(
            RegistroRollo("123", "FJLH78693", "DEP1A14702")
        )
        json_path = self._make_json(tmp_path, registros=[
            {"id_usuario": "123", "id_lote": "FJLH78693",
             "id_ubicacion": "DEP1A14702", "fecha": "2026-07-13 10:00:00"},
        ])
        result = em_con_ubicacion.import_from_android(json_path)
        assert result["reg_dup"] == 1
        assert result["reg_nuevos"] == 0

    def test_importa_inventario_nuevo(self, em_con_ubicacion, tmp_path):
        json_path = self._make_json(tmp_path, inventarios={
            "INV_2026-07": [
                {"id_lote": "FJLH78693", "id_ubicacion": "DEP1A14702",
                 "operario": "123", "fecha": "2026-07-13 10:00:00"},
            ]
        })
        result = em_con_ubicacion.import_from_android(json_path)
        assert result["inv_nuevos"] == 1
        assert "INV_2026-07" in result["inventarios"]

    def test_detecta_duplicados_en_inventario(self, em_con_ubicacion, tmp_path):
        em_con_ubicacion.crear_inventario("INV_2026-07")
        em_con_ubicacion.add_a_inventario("INV_2026-07", "DEP1A14702", "FJLH78693", "123")
        json_path = self._make_json(tmp_path, inventarios={
            "INV_2026-07": [
                {"id_lote": "FJLH78693", "id_ubicacion": "DEP1A14702",
                 "operario": "123", "fecha": "2026-07-13 10:00:00"},
            ]
        })
        result = em_con_ubicacion.import_from_android(json_path)
        assert result["inv_dup"] == 1
        assert result["inv_nuevos"] == 0

    def test_datos_persisten_en_excel(self, em_con_ubicacion, tmp_path):
        """Los datos importados son legibles desde disco después de guardar."""
        json_path = self._make_json(tmp_path, registros=[
            {"id_usuario": "123", "id_lote": "FJLH78693",
             "id_ubicacion": "DEP1A14702", "fecha": "2026-07-13 10:00:00"},
        ])
        em_con_ubicacion.import_from_android(json_path)
        # Verificar con una nueva instancia que lee desde disco
        em2 = ExcelManager(em_con_ubicacion.filepath)
        assert em2.lote_exists("FJLH78693") is True

    def test_no_se_pierde_data_entre_lotes(self, em_con_ubicacion, tmp_path):
        """Multiples registros en un mismo import: todos se guardan."""
        lotes = ["FJLH78693", "ABC123476", "XYZ999988"]
        registros = [
            {"id_usuario": "123", "id_lote": lote,
             "id_ubicacion": "DEP1A14702", "fecha": "2026-07-13 10:00:00"}
            for lote in lotes
        ]
        json_path = self._make_json(tmp_path, registros=registros)
        result = em_con_ubicacion.import_from_android(json_path)
        assert result["reg_nuevos"] == 3
        assert result["reg_dup"] == 0
        for lote in lotes:
            assert em_con_ubicacion.lote_exists(lote) is True

    def test_json_vacio_no_falla(self, em, tmp_path):
        json_path = self._make_json(tmp_path)
        result = em.import_from_android(json_path)
        assert result["reg_nuevos"] == 0
        assert result["inv_nuevos"] == 0

    def test_crea_hoja_inventario_inexistente(self, em_con_ubicacion, tmp_path):
        json_path = self._make_json(tmp_path, inventarios={
            "INV_2026-NUEVO": [
                {"id_lote": "FJLH78693", "id_ubicacion": "DEP1A14702",
                 "operario": "123", "fecha": "2026-07-13 10:00:00"},
            ]
        })
        em_con_ubicacion.import_from_android(json_path)
        assert "INV_2026-NUEVO" in em_con_ubicacion.wb.sheetnames


# ---------------------------------------------------------------------------
# Validación de rollos — integración con checksum_rollo
# ---------------------------------------------------------------------------

class TestValidacionRollos:

    @pytest.mark.parametrize("codigo_9", [
        "FJLH78693", "ABC123476", "XYZ999988",
        "A0B1C2D50", "ZZ00ZZ002", "123456787", "M5N3P7Q97",
    ])
    def test_codigos_validos_aceptados(self, codigo_9):
        valido, msg = validar_lote(codigo_9)
        assert valido is True, f"Debería ser válido: {codigo_9} — {msg}"

    @pytest.mark.parametrize("codigo_invalido,descripcion", [
        ("FJLH78603", "check1 alterado"),
        ("FJLH78694", "check2 alterado"),
        ("FJLH79693", "dígito del producto cambiado"),
        ("JFLH78693", "transposición"),
        ("ABD123476", "letra cambiada"),
        ("ABC123467", "checksum invertido"),
    ])
    def test_codigos_invalidos_rechazados(self, codigo_invalido, descripcion):
        valido, msg = validar_lote(codigo_invalido)
        assert valido is False, f"Debería ser inválido ({descripcion}): {codigo_invalido}"
        assert msg, "El mensaje de error no debe estar vacío"

    def test_longitud_8_rechazada(self):
        valido, msg = validar_lote("FJLH7869")
        assert valido is False
        assert "longitud" in msg.lower() or "Longitud" in msg

    def test_longitud_10_rechazada(self):
        valido, msg = validar_lote("FJLH786930")
        assert valido is False

    def test_caracteres_especiales_rechazados(self):
        valido, msg = validar_lote("FJL-78693")
        assert valido is False
        assert "permitido" in msg.lower()

    def test_checksum_letras_rechazado(self):
        """Los últimos 2 chars deben ser dígitos."""
        valido, msg = validar_lote("FJLH786AB")
        assert valido is False

    def test_roundtrip_calcular_y_validar(self):
        """Cualquier identificador de 7 chars genera un código de 9 que pasa validación."""
        for ident in ["FJLH786", "ABC1234", "ZZ00ZZ0", "1234567"]:
            codigo = ident + calcular_checksum(ident)
            valido, msg = validar_lote(codigo)
            assert valido is True, f"Roundtrip falló para {ident}: {msg}"

    def test_minusculas_aceptadas(self):
        valido, _ = validar_lote("fjlh78693")
        assert valido is True

    def test_espacios_extremos_aceptados(self):
        valido, _ = validar_lote("  FJLH78693  ")
        assert valido is True


# ---------------------------------------------------------------------------
# Validación de ubicaciones — formato 10 chars
# ---------------------------------------------------------------------------

class TestValidacionUbicaciones:

    def test_formato_estandar_tiene_10_chars(self):
        u = Ubicacion("DEP1", "A1", "47", "02")
        assert len(u.id) == 10

    def test_ejemplo_del_prd(self):
        # DEP1A14702: DEP1(4)+A1(2)+47(2)+02(2) = 10 chars
        u = Ubicacion("DEP1", "A1", "47", "02")
        assert u.id == "DEP1A14702"
        assert len(u.id) == 10

    def test_otro_ejemplo_valido(self):
        # D0770000: D077(4)+00(2)+00(2)+00(2) = 10 chars
        u = Ubicacion("D077", "00", "00", "00")
        assert len(u.id) == 10

    def test_deposito_4_chars_pasillo_2_columna_2_nivel_2(self):
        u = Ubicacion("ABCD", "EF", "GH", "IJ")
        assert u.id == "ABCDEFGHIJ"
        assert len(u.id) == 10


# ---------------------------------------------------------------------------
# Credenciales
# ---------------------------------------------------------------------------

class TestCredenciales:

    def test_usuario_123_existe(self):
        assert "123" in CREDENTIALS

    def test_password_correcto(self):
        assert CREDENTIALS["123"] == "Operador1"

    def test_usuario_invalido_retorna_none(self):
        assert CREDENTIALS.get("admin") is None

    def test_password_incorrecto_no_coincide(self):
        assert CREDENTIALS.get("123") != "password_incorrecto"
