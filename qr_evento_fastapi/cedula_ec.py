"""Validación matemática de cédula de ciudadanía ecuatoriana.
No equivale a consultar el Registro Civil.
"""
import re


def cedula_ec_valida(valor: str) -> bool:
    cedula = str(valor or '').strip()
    if not re.fullmatch(r'\d{10}', cedula, flags=re.ASCII):
        return False
    if not 1 <= int(cedula[:2]) <= 24:
        return False
    if int(cedula[2]) >= 6:
        return False
    if len(set(cedula)) == 1:
        return False
    total = 0
    for i, digito in enumerate(cedula[:9]):
        producto = int(digito) * (2 if i % 2 == 0 else 1)
        total += producto - 9 if producto > 9 else producto
    verificador = (10 - total % 10) % 10
    return verificador == int(cedula[-1])


if __name__ == '__main__':
    assert cedula_ec_valida('1719304188')
    assert cedula_ec_valida('0104047568')
    assert not cedula_ec_valida('1719304180')
    assert not cedula_ec_valida('1234567890')
    assert not cedula_ec_valida('171930418a')
    assert not cedula_ec_valida('0000000000')
    print('Pruebas de cedula completadas')
