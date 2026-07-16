"""
Validación de códigos de rollo de tela — Depósito textil.

Formato del código:
    9 caracteres totales
    [0:7]  — identificador del producto (alfanumérico, mayúsculas, ej: "FJLH786")
    [7:9]  — dígitos de verificación (checksum decimal "00"–"99")

Algoritmo (estilo Luhn adaptado para alfanumérico):
    1. Cada caracter se convierte a valor numérico:
           dígito  → int(char)
           letra   → ord(char.upper()) - ord('A') + 10
       Con este esquema: A=10, B=11, …, Z=35; 0=0, …, 9=9.
    2. Pesos alternos sobre los 7 caracteres:
           posiciones 1, 3, 5, 7  →  peso 1
           posiciones 2, 4, 6     →  peso 3
    3. total = suma de (valor × peso) para cada posición
    4. check1 = (10 − (total % 10)) % 10
    5. check2 = (10 − ((total // 10 + check1) % 10)) % 10
    6. Resultado: f"{check1}{check2}"  (siempre 2 dígitos, con cero a la izq.)

El algoritmo detecta:
    - Cualquier dígito cambiado en el identificador o en el checksum.
    - La mayoría de las transposiciones de dos caracteres adyacentes.
"""


def _char_a_valor(c: str) -> int:
    """Convierte un carácter alfanumérico a su valor numérico para el cálculo."""
    c = c.upper()
    if c.isdigit():
        return int(c)
    if c.isalpha():
        return ord(c) - ord('A') + 10  # A=10 … Z=35
    raise ValueError(f"Carácter no permitido: {repr(c)}")


def calcular_checksum(codigo_7: str) -> str:
    """
    Calcula los 2 dígitos de verificación para un identificador de 7 caracteres.

    Args:
        codigo_7: Exactamente 7 caracteres alfanuméricos (A-Z, 0-9).
                  Se acepta en minúsculas; se normaliza a mayúsculas internamente.

    Returns:
        String de 2 dígitos decimales ("00"–"99").

    Raises:
        ValueError: Si la longitud no es 7 o contiene caracteres no permitidos.
    """
    if len(codigo_7) != 7:
        raise ValueError(
            f"El identificador debe tener exactamente 7 caracteres; "
            f"recibido: {len(codigo_7)!r} ({codigo_7!r})"
        )
    codigo_7 = codigo_7.upper()
    for c in codigo_7:
        if not c.isalnum():
            raise ValueError(
                f"Carácter no permitido en posición "
                f"{codigo_7.index(c)}: {repr(c)}"
            )

    pesos = [1, 3, 1, 3, 1, 3, 1]
    total = sum(_char_a_valor(c) * p for c, p in zip(codigo_7, pesos))

    check1 = (10 - (total % 10)) % 10
    check2 = (10 - ((total // 10 + check1) % 10)) % 10
    return f"{check1}{check2}"


def validar_lote(codigo_9: str) -> tuple[bool, str]:
    """
    Valida un código de rollo completo de 9 caracteres.

    Args:
        codigo_9: Código completo de 9 caracteres (7 identificador + 2 checksum).
                  Se acepta en minúsculas; se normaliza a mayúsculas internamente.

    Returns:
        (True, "")  si el código es válido.
        (False, mensaje_de_error)  si es inválido, con descripción del problema.
    """
    if not isinstance(codigo_9, str):
        return False, f"El código debe ser una cadena de texto, recibido: {type(codigo_9).__name__}"

    codigo_9 = codigo_9.strip().upper()

    if len(codigo_9) != 9:
        return False, (
            f"Longitud incorrecta: se esperan 9 caracteres, "
            f"se recibieron {len(codigo_9)}"
        )

    for i, c in enumerate(codigo_9):
        if not c.isalnum():
            return False, f"Carácter no permitido en posición {i + 1}: {repr(c)}"

    # Los últimos 2 deben ser dígitos decimales
    checksum_recibido = codigo_9[7:9]
    if not checksum_recibido.isdigit():
        return False, (
            f"Los últimos 2 caracteres deben ser dígitos decimales (00-99); "
            f"recibido: {repr(checksum_recibido)}"
        )

    identificador = codigo_9[:7]
    checksum_esperado = calcular_checksum(identificador)

    if checksum_recibido != checksum_esperado:
        return False, (
            f"Checksum incorrecto: el código indica '{checksum_recibido}' "
            f"pero el valor calculado es '{checksum_esperado}'"
        )

    return True, ""
