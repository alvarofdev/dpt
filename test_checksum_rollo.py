"""
Tests para checksum_rollo.py

Ejecutar con:
    pytest test_checksum_rollo.py -v

Fixture de códigos válidos (calculados con calcular_checksum):
    FJLH786  →  93  →  FJLH78693
    ABC1234  →  76  →  ABC123476
    XYZ9999  →  88  →  XYZ999988
    A0B1C2D  →  50  →  A0B1C2D50
    ZZ00ZZ0  →  02  →  ZZ00ZZ002
    1234567  →  87  →  123456787
    M5N3P7Q  →  97  →  M5N3P7Q97
"""

import pytest
from checksum_rollo import calcular_checksum, validar_lote


# ---------------------------------------------------------------------------
# Fixtures — valores calculados con el algoritmo, usados como referencia
# ---------------------------------------------------------------------------

CODIGOS_VALIDOS = [
    # (identificador_7, checksum_2, codigo_9)
    ("FJLH786", "93", "FJLH78693"),
    ("ABC1234", "76", "ABC123476"),
    ("XYZ9999", "88", "XYZ999988"),
    ("A0B1C2D", "50", "A0B1C2D50"),
    ("ZZ00ZZ0", "02", "ZZ00ZZ002"),
    ("1234567", "87", "123456787"),
    ("M5N3P7Q", "97", "M5N3P7Q97"),
]

# Códigos con checksum o identificador alterado — todos deben fallar
CODIGOS_INVALIDOS = [
    # (codigo_9_incorrecto, descripcion_del_error_inyectado)
    ("FJLH78603", "check1 incrementado en 1 (9→0)"),
    ("FJLH78694", "check2 incrementado en 1 (3→4)"),
    ("FJLH79693", "dígito del producto cambiado (8→9 en pos 6)"),
    ("JFLH78693", "primeros dos chars transpuestos (FJ→JF)"),
    ("ABD123476", "letra del producto cambiada (C→D en pos 3)"),
    ("ABC123467", "dígitos del checksum invertidos (76→67)"),
]


# ---------------------------------------------------------------------------
# Tests de calcular_checksum
# ---------------------------------------------------------------------------

class TestCalcularChecksum:

    @pytest.mark.parametrize("identificador,checksum_esperado,_", CODIGOS_VALIDOS)
    def test_checksum_correcto(self, identificador, checksum_esperado, _):
        """El checksum calculado coincide con el valor de referencia."""
        assert calcular_checksum(identificador) == checksum_esperado

    def test_retorna_siempre_dos_digitos(self):
        """El resultado tiene siempre exactamente 2 caracteres."""
        for identificador, _, _ in CODIGOS_VALIDOS:
            resultado = calcular_checksum(identificador)
            assert len(resultado) == 2, f"Largo incorrecto para {identificador}: {resultado!r}"
            assert resultado.isdigit(), f"No son dígitos para {identificador}: {resultado!r}"

    def test_checksum_con_cero_a_izquierda(self):
        """ZZ00ZZ0 produce '02': verifica que el cero inicial se conserva."""
        assert calcular_checksum("ZZ00ZZ0") == "02"

    def test_acepta_minusculas(self):
        """La función normaliza a mayúsculas antes de calcular."""
        assert calcular_checksum("fjlh786") == calcular_checksum("FJLH786")
        assert calcular_checksum("abc1234") == calcular_checksum("ABC1234")

    def test_es_determinista(self):
        """Dos llamadas con el mismo input producen el mismo output."""
        for identificador, _, _ in CODIGOS_VALIDOS:
            assert calcular_checksum(identificador) == calcular_checksum(identificador)

    def test_error_longitud_corta(self):
        """Menos de 7 caracteres lanza ValueError."""
        with pytest.raises(ValueError, match="7 caracteres"):
            calcular_checksum("FJLH78")

    def test_error_longitud_larga(self):
        """Más de 7 caracteres lanza ValueError."""
        with pytest.raises(ValueError, match="7 caracteres"):
            calcular_checksum("FJLH7860")

    def test_error_caracter_especial(self):
        """Caracteres no alfanuméricos lanzan ValueError."""
        with pytest.raises(ValueError):
            calcular_checksum("FJL-786")

    def test_error_espacio(self):
        """Espacios internos lanzan ValueError."""
        with pytest.raises(ValueError):
            calcular_checksum("FJL 786")

    def test_sensible_a_un_cambio(self):
        """Cambiar un solo carácter del identificador cambia el checksum."""
        base = calcular_checksum("FJLH786")
        variante = calcular_checksum("FJLH787")   # último dígito 6→7
        assert base != variante


# ---------------------------------------------------------------------------
# Tests de validar_lote
# ---------------------------------------------------------------------------

class TestValidarLote:

    @pytest.mark.parametrize("_id,_cs,codigo_9", CODIGOS_VALIDOS)
    def test_codigos_validos(self, _id, _cs, codigo_9):
        """Cada código de la fixture válida pasa la validación."""
        valido, mensaje = validar_lote(codigo_9)
        assert valido is True
        assert mensaje == ""

    @pytest.mark.parametrize("codigo_invalido,descripcion", CODIGOS_INVALIDOS)
    def test_codigos_invalidos(self, codigo_invalido, descripcion):
        """Cada código alterado debe ser rechazado."""
        valido, mensaje = validar_lote(codigo_invalido)
        assert valido is False, f"Debería ser inválido ({descripcion}): {codigo_invalido}"
        assert isinstance(mensaje, str) and len(mensaje) > 0, \
            "El mensaje de error no debe estar vacío"

    def test_acepta_minusculas(self):
        """validar_lote normaliza a mayúsculas antes de validar."""
        valido, _ = validar_lote("fjlh78693")
        assert valido is True

    def test_acepta_espacios_extremos(self):
        """Strip de espacios al inicio/fin antes de validar."""
        valido, _ = validar_lote("  FJLH78693  ")
        assert valido is True

    def test_longitud_incorrecta_corta(self):
        """Código con menos de 9 caracteres es inválido."""
        valido, mensaje = validar_lote("FJLH7869")
        assert valido is False
        assert "Longitud" in mensaje or "longitud" in mensaje

    def test_longitud_incorrecta_larga(self):
        """Código con más de 9 caracteres es inválido."""
        valido, mensaje = validar_lote("FJLH786930")
        assert valido is False
        assert "Longitud" in mensaje or "longitud" in mensaje

    def test_checksum_con_letras_en_posicion_78(self):
        """Los últimos 2 caracteres deben ser dígitos decimales, no letras."""
        valido, mensaje = validar_lote("FJLH786AB")
        assert valido is False
        assert "dígitos decimales" in mensaje or "dígito" in mensaje.lower()

    def test_caracter_especial_en_identificador(self):
        """Un guion u otro símbolo en el identificador es inválido."""
        valido, mensaje = validar_lote("FJL-78693")
        assert valido is False
        assert "permitido" in mensaje.lower()

    def test_tipo_no_string(self):
        """Pasar un entero devuelve (False, mensaje) sin lanzar excepción."""
        valido, mensaje = validar_lote(123456789)  # type: ignore[arg-type]
        assert valido is False
        assert isinstance(mensaje, str)

    def test_mensaje_indica_checksum_esperado(self):
        """El mensaje de error de checksum muestra el valor correcto esperado."""
        # FJLH78693 es válido; FJLH78699 tiene el checksum '99' que es incorrecto
        valido, mensaje = validar_lote("FJLH78699")
        assert valido is False
        # El mensaje debe mencionar el checksum correcto '93'
        assert "93" in mensaje

    def test_codigo_todo_ceros_invalido(self):
        """000000000 → checksum de 0000000 es '00', así que 000000000 es válido."""
        cs = calcular_checksum("0000000")
        if cs == "00":
            valido, _ = validar_lote("000000000")
            assert valido is True
        else:
            valido, _ = validar_lote("000000000")
            assert valido is False

    def test_roundtrip_todos_los_fixtures(self):
        """Para cada identificador de la fixture: calcular_checksum + validar_lote = True."""
        identificadores = [
            "FJLH786", "ABC1234", "XYZ9999",
            "A0B1C2D", "ZZ00ZZ0", "1234567", "M5N3P7Q",
        ]
        for ident in identificadores:
            codigo = ident + calcular_checksum(ident)
            valido, msg = validar_lote(codigo)
            assert valido is True, f"Roundtrip fallido para {ident}: {msg}"
